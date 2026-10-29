#!/usr/bin/env node
/**
 * PDF 工具箱 · CLI（Node ESM）
 *
 * 设计：Node 只做参数解析 / 依赖探测 / 输出命名与展示；
 *       真正的 PDF 读写交给 Python + PyMuPDF（lib/pdf_ops.py），
 *       通过 spawnSync 调用，stdin 传 JSON 任务，stdout 收 JSON 结果。
 *
 * 契约：
 *   - 不改输入文件；只写 --out（默认 <输入目录>/_输出）
 *   - 同名不覆盖，自动加 -1 / -2 后缀
 *   - 依赖缺失 / 参数错误明确报「缺什么、怎么装」，非 0 退出
 */
import { spawnSync } from 'node:child_process';
import {
  existsSync, mkdirSync, mkdtempSync, readdirSync, statSync,
} from 'node:fs';
import { basename, dirname, extname, join, resolve } from 'node:path';
import { tmpdir } from 'node:os';
import { fileURLToPath } from 'node:url';

const PKG_ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const OPS = join(PKG_ROOT, 'lib', 'pdf_ops.py');
const CMD_NAME = 'dsh-pdf-kit';

const COMMANDS = {
  merge: { actions: [], modes: {} },
  split: {
    actions: ['page', 'by-page', 'every', 'every-n'],
    modes: { page: 'page', 'by-page': 'page', every: 'every', 'every-n': 'every' },
  },
  compress: { actions: [], modes: {} },
  watermark: { actions: ['text'], modes: { text: 'text' } },
  'extract-text': { actions: [], modes: {} },
  'extract-images': { actions: [], modes: {} },
};

const VALUE_FLAGS = new Set([
  '--input', '--out', '--n', '--quality', '--dpi', '--text',
  '--fontsize', '--opacity', '--min-bytes',
]);
const BOOL_FLAGS = new Set(['--json', '--demo', '--probe', '--help', '-h']);

// ------------------------------------------------------------------ 输出/退出

function emit(json, payload) {
  if (json) console.log(JSON.stringify(payload));
}

function fail(message, { hint = '', code = 2, json = false, extra = {} } = {}) {
  if (json) {
    console.log(JSON.stringify({ ok: false, error: message, hint, ...extra }));
  } else {
    console.error(`❌ ${message}`);
    if (hint) {
      for (const line of String(hint).split('\n')) console.error(`   ${line}`);
    }
  }
  process.exit(code);
}

function fmtBytes(n) {
  if (n == null) return '—';
  const units = ['B', 'KB', 'MB', 'GB'];
  let v = n;
  let i = 0;
  while (v >= 1024 && i < units.length - 1) { v /= 1024; i += 1; }
  return `${v >= 100 || i === 0 ? Math.round(v) : v.toFixed(1)} ${units[i]}`;
}

// ------------------------------------------------------------------ 帮助

function printHelp() {
  console.log(`PDF 工具箱 · ${CMD_NAME}

一句话：PDF 合并 / 按页拆分 / 压缩 / 整页文字水印 / 提取文本 / 导出图片

用法：
  ${CMD_NAME} <子命令> [动作] --input <文件或目录> [--out <目录>] [--json]
  ${CMD_NAME} --demo
  ${CMD_NAME} --help

子命令（动作可省略；也支持 ${CMD_NAME} pdf merge 这种带前缀写法）：
  merge           合并多个 PDF（--input 可重复，或逗号分隔）
  split           拆分：默认按页；动作 every-n / 选项 --n N 表示每 N 页一份
  compress        重存 + 图片降采样重编码（--quality 1-95，--dpi 36-600）
  watermark       整页平铺文字水印（--text 必填，--fontsize，--opacity 0-1）
  extract-text    每页文本各导出一个 .txt（UTF-8）
  extract-images  导出页面内嵌图片（保留原始字节）

示例：
  ${CMD_NAME} merge --input a.pdf,b.pdf --out out
  ${CMD_NAME} split every-n --n 2 --input report.pdf
  ${CMD_NAME} watermark --text "内部资料" --input a.pdf
  ${CMD_NAME} compress --quality 60 --input big.pdf
  ${CMD_NAME} --demo            # 自造样例跑完整流程，打印产物路径

通用选项：
  --input <路径>   输入 PDF 或目录（可重复）
  --out <目录>     输出目录（默认 <输入目录>/_输出）
  --json           机器可读输出
  --probe          探测 Python / PyMuPDF 版本（诊断用）
  -h, --help       显示帮助

约定：
  · 不修改输入文件；只写输出目录；同名文件自动加 -1 / -2，不覆盖。
  · 依赖 python3 + PyMuPDF；缺失时明确提示「缺什么、怎么装」。
  · 退出码：0 成功 / 2 参数或输入错误 / 3 依赖缺失 / 4 执行失败。`);
}

// ------------------------------------------------------------------ 参数解析

function parseArgs(argv) {
  const opts = {
    inputs: [], out: null, json: false, demo: false, probe: false, help: false,
    n: null, quality: null, dpi: null, text: null, fontsize: null, opacity: null,
    minBytes: null, positionals: [],
  };
  for (let i = 0; i < argv.length; i += 1) {
    const raw = argv[i];
    const eq = raw.startsWith('--') ? raw.indexOf('=') : -1;
    const flag = eq > 0 ? raw.slice(0, eq) : raw;
    const inline = eq > 0 ? raw.slice(eq + 1) : null;
    if (BOOL_FLAGS.has(flag)) {
      if (flag === '--json') opts.json = true;
      else if (flag === '--demo') opts.demo = true;
      else if (flag === '--probe') opts.probe = true;
      else opts.help = true;
      continue;
    }
    if (VALUE_FLAGS.has(flag)) {
      let value = inline;
      if (value == null) {
        value = argv[i + 1];
        if (value == null || (value.startsWith('--') && value !== '-')) {
          fail(`参数 ${flag} 缺少值`, { hint: `例如：${flag} <值>`, json: opts.json });
        }
        i += 1;
      }
      switch (flag) {
        case '--input':
          for (const part of value.split(',')) {
            const t = part.trim();
            if (t) opts.inputs.push(t);
          }
          break;
        case '--out': opts.out = value; break;
        case '--n': opts.n = value; break;
        case '--quality': opts.quality = value; break;
        case '--dpi': opts.dpi = value; break;
        case '--text': opts.text = value; break;
        case '--fontsize': opts.fontsize = value; break;
        case '--opacity': opts.opacity = value; break;
        case '--min-bytes': opts.minBytes = value; break;
        default: break;
      }
      continue;
    }
    if (raw.startsWith('-') && raw !== '-') {
      fail(`未知参数：${raw}`, { hint: '用 --help 查看支持的参数', json: opts.json });
    }
    opts.positionals.push(raw);
  }
  return opts;
}

// ------------------------------------------------------------------ 依赖探测

function findPython() {
  const candidates = [
    process.env.DSH_PYTHON,
    '/opt/homebrew/bin/python3',
    '/usr/local/bin/python3',
    'python3',
  ].filter(Boolean);
  const tried = [];
  for (const cand of candidates) {
    const r = spawnSync(cand, ['-c', 'import sys; print(sys.executable)'], { encoding: 'utf8' });
    tried.push(cand);
    if (!r.error && r.status === 0) {
      const exe = (r.stdout || '').trim().split('\n').filter(Boolean).pop();
      if (exe) return { python: exe, tried };
    }
  }
  return { python: null, tried };
}

function ensurePython(json) {
  const { python, tried } = findPython();
  if (!python) {
    fail('缺少 python3：找不到可用的 Python 解释器', {
      code: 3,
      json,
      hint: `已尝试：${tried.join('、')}
安装：brew install python3
或设置环境变量 DSH_PYTHON 指向 python3 绝对路径`,
      extra: { dependency: 'python3' },
    });
  }
  return python;
}

function ensureFitz(python, json) {
  const code = 'import fitz, json; print(json.dumps({"fitz": getattr(fitz, "VersionBind", "unknown")}))';
  const r = spawnSync(python, ['-c', code], { encoding: 'utf8' });
  if (r.error || r.status !== 0) {
    const detail = (r.stderr || r.stdout || String(r.error || '')).trim().split('\n').slice(-3).join('\n');
    fail('缺少 PyMuPDF（python 包 fitz）：无法读写 PDF', {
      code: 3,
      json,
      hint: `安装：${python} -m pip install PyMuPDF
如系统 Python 受限，可用：${python} -m pip install --user PyMuPDF
${detail ? `底层报错：${detail}` : ''}`,
      extra: { dependency: 'PyMuPDF' },
    });
  }
  const line = (r.stdout || '').trim().split('\n').filter(Boolean).pop();
  try {
    return JSON.parse(line);
  } catch {
    return { fitz: 'unknown' };
  }
}

function runOp(python, job) {
  const r = spawnSync(python, [OPS], {
    input: JSON.stringify(job),
    encoding: 'utf8',
    maxBuffer: 64 * 1024 * 1024,
  });
  const stdout = (r.stdout || '').trim();
  const stderr = (r.stderr || '').trim();
  let parsed = null;
  if (stdout) {
    const lines = stdout.split('\n').filter(Boolean);
    for (let i = lines.length - 1; i >= 0; i -= 1) {
      try { parsed = JSON.parse(lines[i]); break; } catch { /* 继续找 */ }
    }
  }
  if (parsed) return parsed;
  return {
    ok: false,
    error: 'python-op-failed',
    detail: `Python 侧没有返回结果（退出码 ${r.status}）`,
    stderr: stderr.split('\n').slice(-6).join('\n'),
  };
}

function unwrap(result, step, json) {
  if (!result || result.ok !== true) {
    const detail = result?.detail || result?.error || '未知错误';
    const stderr = result?.stderr ? `\n${result.stderr}` : '';
    fail(`${step} 执行失败：${detail}${stderr}`, {
      code: 4,
      json,
      hint: result?.hint || '可加 --json 查看结构化错误；输入文件损坏或加密时也会失败',
      extra: { step, detail: String(detail) },
    });
  }
  return result;
}

// ------------------------------------------------------------------ 输入/输出

const PDF_EXT = '.pdf';

function collectInputs(opts, json) {
  const seen = new Set();
  const files = [];
  for (const raw of opts.inputs) {
    const abs = resolve(raw);
    if (!existsSync(abs)) {
      fail(`输入不存在：${raw}（解析为 ${abs}）`, {
        code: 2,
        json,
        hint: '确认路径拼写；支持 PDF 文件或包含 PDF 的目录',
        extra: { input: raw },
      });
    }
    const st = statSync(abs);
    if (st.isDirectory()) {
      const entries = readdirSync(abs)
        .filter((f) => extname(f).toLowerCase() === PDF_EXT)
        .filter((f) => {
          const full = join(abs, f);
          try { return statSync(full).isFile(); } catch { return false; }
        })
        .sort((a, b) => a.localeCompare(b, 'zh-CN'));
      if (entries.length === 0) {
        fail(`目录里没有 PDF 文件：${raw}`, {
          code: 2, json, hint: '把 PDF 放进该目录，或直接传文件路径', extra: { input: raw },
        });
      }
      for (const f of entries) {
        const full = resolve(join(abs, f));
        if (!seen.has(full)) { seen.add(full); files.push(full); }
      }
    } else if (st.isFile()) {
      if (extname(abs).toLowerCase() !== PDF_EXT) {
        fail(`不是 PDF 文件：${raw}`, {
          code: 2, json, hint: '本工具只处理 .pdf；其它格式请先转换', extra: { input: raw },
        });
      }
      if (!seen.has(abs)) { seen.add(abs); files.push(abs); }
    } else {
      fail(`无法识别的输入（既不是文件也不是目录）：${raw}`, { code: 2, json });
    }
  }
  return files;
}

function defaultOutDir(files, opts) {
  if (opts.out) return resolve(opts.out);
  const firstDir = dirname(files[0]);
  return join(firstDir, '_输出');
}

function ensureDir(dir, json) {
  try {
    mkdirSync(dir, { recursive: true });
  } catch (e) {
    fail(`无法创建输出目录：${dir}`, {
      code: 2, json, hint: String(e?.message || e), extra: { out: dir },
    });
  }
}

function stemOf(file) {
  return basename(file, extname(file));
}

// ------------------------------------------------------------------ 人类输出

function printResult(step, result) {
  const s = result.stats || {};
  const outputs = result.outputs || [];
  if (step === 'merge') {
    console.log(`✅ merge 完成 · ${s.files ?? '?'} 个 PDF → ${s.pages} 页`);
  } else if (step === 'split') {
    console.log(`✅ split 完成 · 共 ${s.pages} 页 → ${s.parts} 个文件（${s.mode === 'every' ? `每 ${s.n} 页` : '按页'}）`);
  } else if (step === 'compress') {
    const pct = s.beforeBytes ? ((1 - s.afterBytes / s.beforeBytes) * 100) : 0;
    const mark = s.smaller ? `-${pct.toFixed(1)}%` : `未变小（${pct.toFixed(1)}%）`;
    console.log(`✅ compress 完成 · ${fmtBytes(s.beforeBytes)} → ${fmtBytes(s.afterBytes)}（${mark}）· 图片 ${s.images ?? 0} 张`);
  } else if (step === 'watermark') {
    console.log(`✅ watermark 完成 · ${s.pages} 页 · 文字「${s.text}」平铺 ${s.placements} 处 · 不透明度 ${s.opacity}`);
  } else if (step === 'extract-text') {
    console.log(`✅ extract-text 完成 · ${s.pages} 页 → ${s.chars} 字符`);
  } else if (step === 'extract-images') {
    console.log(`✅ extract-images 完成 · ${s.pages} 页 → ${s.images} 张图片${s.skipped ? `（跳过 ${s.skipped} 张）` : ''}`);
  } else {
    console.log(`✅ ${step} 完成`);
  }
  for (const out of outputs) console.log(`   产物：${out}`);
}

// ------------------------------------------------------------------ 演示

function runDemo(opts) {
  const json = opts.json;
  const python = ensurePython(json);
  const fitzInfo = ensureFitz(python, json);
  const work = mkdtempSync(join(tmpdir(), 'dsh-pdf-kit-demo-'));
  const outRoot = join(work, 'output');
  const make = unwrap(runOp(python, { action: 'make-demo', dir: work }), 'make-demo', json);
  const { textA, textB, heavy } = make.stats;
  const steps = [];

  const run = (step, job) => {
    const res = unwrap(runOp(python, job), step, json);
    steps.push({ step, outputs: res.outputs, stats: res.stats });
    if (!json) printResult(step, res);
    return res;
  };

  if (!json) {
    console.log('');
    console.log('  PDF 工具箱 · 演示模式（自造样例，不碰你的文件）');
    console.log(`  Python：${python} · PyMuPDF ${fitzInfo.fitz}`);
    console.log(`  工作目录：${work}`);
    console.log('');
  }

  run('merge', { action: 'merge', files: [textA, textB], outDir: join(outRoot, 'merge'), base: 'demo' });
  run('split', { action: 'split', file: textA, outDir: join(outRoot, 'split-page'), base: 'demo-a', mode: 'page' });
  run('split', { action: 'split', file: textA, outDir: join(outRoot, 'split-every'), base: 'demo-a', mode: 'every', n: 2 });
  run('compress', { action: 'compress', file: heavy, outDir: join(outRoot, 'compress'), base: 'demo-heavy', quality: 60, dpi: 72 });
  run('watermark', { action: 'watermark', file: textA, outDir: join(outRoot, 'watermark'), base: 'demo-a', text: 'DSH 深鲸示例水印', fontsize: 28, opacity: 0.16 });
  run('extract-text', { action: 'extract-text', file: textA, outDir: join(outRoot, 'text'), base: 'demo-a' });
  run('extract-images', { action: 'extract-images', file: textA, outDir: join(outRoot, 'images'), base: 'demo-a' });

  if (json) {
    emit(true, { ok: true, demo: true, workdir: work, python, fitz: fitzInfo.fitz, steps });
    return;
  }
  console.log('');
  console.log(`✅ 演示完成：7 个动作全部真跑通，产物共 ${steps.reduce((n, s) => n + s.outputs.length, 0)} 个`);
  console.log(`   查看：${outRoot}`);
  console.log('   说明：演示文件在临时目录，可自行删除。');
}

// ------------------------------------------------------------------ 主流程

function main() {
  const argv = process.argv.slice(2);
  if (argv.length === 0) {
    printHelp();
    process.exit(0);
  }
  const opts = parseArgs(argv);

  if (opts.help) {
    printHelp();
    process.exit(0);
  }

  if (opts.probe) {
    const python = ensurePython(opts.json);
    const info = ensureFitz(python, opts.json);
    if (opts.json) {
      emit(true, { ok: true, python, fitz: info.fitz });
    } else {
      console.log('');
      console.log('  PDF 工具箱 · 本机能力探测');
      console.log(`  python : ${python}`);
      console.log(`  PyMuPDF: ${info.fitz}`);
      console.log('');
    }
    process.exit(0);
  }

  if (opts.demo) {
    runDemo(opts);
    return;
  }

  // 子命令与动作
  const positionals = [...opts.positionals];
  if (positionals[0] === 'pdf') positionals.shift();
  const cmd = positionals.shift();
  if (!cmd) {
    fail('缺少子命令', {
      code: 2,
      json: opts.json,
      hint: `可用子命令：${Object.keys(COMMANDS).join(' / ')}\n用 ${CMD_NAME} --help 查看完整用法`,
    });
  }
  const spec = COMMANDS[cmd];
  if (!spec) {
    fail(`未知子命令：${cmd}`, {
      code: 2,
      json: opts.json,
      hint: `可用子命令：${Object.keys(COMMANDS).join(' / ')}\n用 ${CMD_NAME} --help 查看完整用法`,
      extra: { command: cmd },
    });
  }
  let action = null;
  if (positionals.length > 0 && spec.actions.includes(positionals[0])) {
    action = positionals.shift();
  }
  // 其余位置参数当作输入（方便 `merge a.pdf b.pdf`）
  for (const p of positionals) opts.inputs.push(p);

  if (opts.inputs.length === 0) {
    fail('缺少 --input（或直接给出 PDF 路径）', {
      code: 2,
      json: opts.json,
      hint: `例如：${CMD_NAME} ${cmd} --input <文件或目录>\n或先跑 ${CMD_NAME} --demo 看效果`,
    });
  }

  // 选项校验
  const intOpt = (name, raw, min, max, dflt) => {
    if (raw == null) return dflt;
    const v = Number(raw);
    if (!Number.isInteger(v) || v < min || v > max) {
      fail(`--${name} 需要 ${min}~${max} 的整数，收到：${raw}`, { code: 2, json: opts.json });
    }
    return v;
  };
  const floatOpt = (name, raw, min, max, dflt) => {
    if (raw == null) return dflt;
    const v = Number(raw);
    if (!Number.isFinite(v) || v < min || v > max) {
      fail(`--${name} 需要 ${min}~${max} 的数值，收到：${raw}`, { code: 2, json: opts.json });
    }
    return v;
  };

  let mode = 'page';
  let n = null;
  if (cmd === 'split') {
    if (action) mode = spec.modes[action] || 'page';
    if (opts.n != null) { mode = 'every'; n = intOpt('n', opts.n, 1, 100000, 1); }
    if (mode === 'every' && n == null) n = 2;
  }
  if (cmd === 'watermark' && !opts.text) {
    fail('watermark 缺少 --text（水印文字）', {
      code: 2,
      json: opts.json,
      hint: '例如：dsh-pdf-kit watermark --text "内部资料" --input a.pdf',
    });
  }

  const quality = intOpt('quality', opts.quality, 1, 95, 65);
  const dpi = intOpt('dpi', opts.dpi, 36, 600, 96);
  const fontsize = floatOpt('fontsize', opts.fontsize, 6, 200, 30);
  const opacity = floatOpt('opacity', opts.opacity, 0.01, 1, 0.16);
  const minBytes = intOpt('min-bytes', opts.minBytes, 0, 1e9, 0);

  const files = collectInputs(opts, opts.json);
  if (cmd === 'merge' && files.length < 2) {
    fail(`merge 至少需要 2 个 PDF，当前只有 ${files.length} 个`, {
      code: 2,
      json: opts.json,
      hint: '用 --input a.pdf,b.pdf 或把多个 PDF 放进同一个目录后传目录',
      extra: { files },
    });
  }

  const python = ensurePython(opts.json);
  const fitzInfo = ensureFitz(python, opts.json);
  const outDir = defaultOutDir(files, opts);
  ensureDir(outDir, opts.json);

  if (!opts.json) {
    console.log('');
    console.log(`  PDF 工具箱 · ${cmd}${mode === 'every' && cmd === 'split' ? `（每 ${n} 页）` : ''}`);
    console.log(`  输入：${files.length} 个文件`);
    console.log(`  输出：${outDir}`);
    console.log(`  环境：${python} · PyMuPDF ${fitzInfo.fitz}`);
    console.log('');
  }

  const results = [];
  if (cmd === 'merge') {
    const job = {
      action: 'merge', files, outDir, base: stemOf(files[0]),
    };
    const res = unwrap(runOp(python, job), 'merge', opts.json);
    results.push({ step: 'merge', inputs: files, ...res });
    if (!opts.json) printResult('merge', res);
  } else {
    for (const file of files) {
      let job;
      if (cmd === 'split') {
        job = { action: 'split', file, outDir, base: stemOf(file), mode, n };
      } else if (cmd === 'compress') {
        job = { action: 'compress', file, outDir, base: stemOf(file), quality, dpi };
      } else if (cmd === 'watermark') {
        job = { action: 'watermark', file, outDir, base: stemOf(file), text: opts.text, fontsize, opacity };
      } else if (cmd === 'extract-text') {
        job = { action: 'extract-text', file, outDir, base: stemOf(file) };
      } else if (cmd === 'extract-images') {
        job = { action: 'extract-images', file, outDir, base: stemOf(file), minBytes };
      }
      const res = unwrap(runOp(python, job), cmd, opts.json);
      results.push({ step: cmd, input: file, ...res });
      if (!opts.json) printResult(cmd, res);
    }
  }

  if (opts.json) {
    const outputs = results.flatMap((r) => r.outputs || []);
    console.log(JSON.stringify({
      ok: true,
      command: cmd,
      action,
      outDir,
      files: files.length,
      count: outputs.length,
      outputs,
      results,
    }));
  } else {
    console.log('');
    console.log(`✅ 全部完成：${results.reduce((acc, r) => acc + (r.outputs?.length || 0), 0)} 个产物 · 输入未改动`);
  }
}

main();
