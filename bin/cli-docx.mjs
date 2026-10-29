#!/usr/bin/env node
/**
 * Word 批量处理 CLI
 *
 * 形态：
 *   dsh-docx-batch --help                 （无参数也打印 help，退出 0）
 *   dsh-docx-batch --demo                 （自造样例跑完整流程，退出 0）
 *   dsh-docx-batch <action> --input <文件或目录> [--out <目录>] [--json]
 *
 * 真实逻辑全部委托 python3 + python-docx（lib/docx_tool.py），
 * Node 只负责参数校验、调用与人类可读输出。
 */
import { spawnSync } from 'node:child_process';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const PKG_ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const PY_SCRIPT = join(PKG_ROOT, 'lib', 'docx_tool.py');
const TITLE = 'Word 批量处理';
const CMD = 'dsh-docx-batch';
const ACTIONS = {
  replace: '批量替换文本（--from 旧 --to 新，可多组）',
  extract: '提取正文文本到 .txt',
  merge: '按文件名顺序合并成一个 docx',
  headers: '列出页眉页脚文本',
  'to-md': '段落/表格转 Markdown',
};
const PY_HINT = [
  '需要 python3 + python-docx。安装：',
  '  python3 -m pip install python-docx',
  '没有 python3 时：',
  '  macOS: brew install python3',
  '  Ubuntu: sudo apt install python3',
].join('\n');
const LINE = '─'.repeat(56);

function help() {
  console.log(`${TITLE}

用法：
  ${CMD} <action> --input <文件或目录> [--out <目录>] [--json]
  ${CMD} --demo
  ${CMD} --help

子命令：
${Object.entries(ACTIONS).map(([k, v]) => `  ${k.padEnd(9)}${v}`).join('\n')}

参数：
  --input, -i   .docx 文件或目录（目录递归处理）
  --out, -o     输出目录；默认 <输入目录>/_输出；同名文件不覆盖，自动加序号
  --from        要替换的旧文本（可重复，与 --to 按顺序配对）
  --to          替换后的新文本（可重复）
  --json        输出 JSON（给脚本/CI 用）
  --demo        自造样例文件，跑完整流程并打印产物路径

示例：
  ${CMD} replace --input ./合同 --from 深鲸科技 --to 深海鲸
  ${CMD} extract --input ./合同 --out ./文本
  ${CMD} merge --input ./周报 --out ./合并
  ${CMD} headers --input ./合同
  ${CMD} to-md --input ./合同 --out ./md

依赖：python3 + python-docx（缺失时会提示安装命令）`);
  process.exit(0);
}

function fail(message, hint = '', code = 2, json = false) {
  if (json) {
    console.log(JSON.stringify({ ok: false, code, message, hint }, null, 2));
  } else {
    console.error(`❌ ${message}`);
    if (hint) console.error(`\n${hint}`);
  }
  process.exit(code);
}

function findPython() {
  const candidates = [
    process.env.DSH_PYTHON,
    process.env.DOCX_BATCH_PYTHON,
    'python3',
    '/opt/homebrew/bin/python3',
    '/usr/local/bin/python3',
    '/usr/bin/python3',
  ].filter(Boolean);
  for (const cmd of candidates) {
    const r = spawnSync(cmd, ['-c', 'import sys; print(sys.version_info[0])'], {
      encoding: 'utf8',
      timeout: 10000,
    });
    if (!r.error && r.status === 0) return cmd;
  }
  return null;
}

function parseArgs(argv) {
  const opt = { action: null, input: null, out: null, from: [], to: [], json: false, demo: false };
  for (let i = 0; i < argv.length; i++) {
    const raw = argv[i];
    const eq = raw.indexOf('=');
    const flag = eq > 0 ? raw.slice(0, eq) : raw;
    const inline = eq > 0 ? raw.slice(eq + 1) : undefined;
    if (flag === '--json') {
      opt.json = true;
      continue;
    }
    if (flag === '--demo') {
      opt.demo = true;
      continue;
    }
    if (flag === '--help' || flag === '-h') help();
    if (flag === '--input' || flag === '-i' || flag === '--out' || flag === '-o' ||
        flag === '--from' || flag === '--to') {
      let value = inline;
      if (value === undefined) {
        value = argv[i + 1];
        if (value === undefined || value.startsWith('--')) {
          fail(`参数 ${flag} 缺少取值`, `例：${flag} <值>`);
        }
        i += 1;
      }
      if (flag === '--input' || flag === '-i') opt.input = value;
      else if (flag === '--out' || flag === '-o') opt.out = value;
      else if (flag === '--from') opt.from.push(value);
      else opt.to.push(value);
      continue;
    }
    if (flag.startsWith('-')) {
      fail(`未知参数：${raw}`, `用 ${CMD} --help 查看支持的参数。`);
    }
    if (!opt.action) {
      opt.action = raw;
      continue;
    }
    fail(`多余的参数：${raw}`, '一次只能执行一个子命令。');
  }
  return opt;
}

function human(res) {
  console.log('');
  console.log(`  ${TITLE} · ${res.action}`);
  console.log(`  ${LINE}`);
  if (res.input) console.log(`  输入       ${res.input}`);
  if (res.out) console.log(`  输出目录   ${res.out}`);
  for (const f of res.files || []) {
    let extra = '';
    if (f.replacements !== undefined) extra = ` · 替换 ${f.replacements} 处`;
    else if (f.chars !== undefined) extra = ` · ${f.chars} 字符`;
    else if (f.sections !== undefined) extra = ` · ${f.sections} 节`;
    else if (f.paragraphs !== undefined) extra = ` · ${f.paragraphs} 段`;
    console.log(`  ✓ ${f.input_name} → ${f.output_name}${extra}`);
  }
  if (res.summary) console.log(`\n  ${res.summary}`);
  if (res.artifacts && res.artifacts.length) {
    console.log('\n  产物：');
    for (const a of res.artifacts) console.log(`    ${a}`);
  }
  console.log('');
}

function humanDemo(res) {
  console.log('');
  console.log(`  ${TITLE} · 演示（--demo）`);
  console.log(`  ${LINE}`);
  console.log(`  演示目录   ${res.demoRoot}`);
  console.log('');
  for (const s of res.steps || []) {
    console.log(`  ✓ ${String(s.action).padEnd(8)} ${s.summary}`);
  }
  console.log('\n  产物：');
  for (const a of res.artifacts || []) console.log(`    ${a}`);
  console.log('');
}

function callPython(py, action, payload) {
  const r = spawnSync(py, [PY_SCRIPT, action, JSON.stringify(payload)], {
    encoding: 'utf8',
    maxBuffer: 128 * 1024 * 1024,
  });
  if (r.error) {
    fail(`调用 python 失败：${r.error.message}`, PY_HINT, 3);
  }
  const lines = (r.stdout || '').trim().split('\n').filter(Boolean);
  let result = null;
  try {
    result = JSON.parse(lines[lines.length - 1] || '');
  } catch {
    /* 下面统一报错 */
  }
  if (!result) {
    const detail = (r.stderr || '').trim().split('\n').slice(-5).join('\n');
    fail('Python 处理失败（没有拿到有效结果）', detail || PY_HINT, r.status || 1);
  }
  return result;
}

const argv = process.argv.slice(2);
if (argv.length === 0) help();
if (argv.includes('--help') || argv.includes('-h')) help();

const opt = parseArgs(argv);
const python = findPython();
if (!python) {
  fail('未找到 python3，无法处理 Word 文件', PY_HINT, 3, opt.json);
}

if (opt.demo) {
  const res = callPython(python, 'demo', {});
  if (!res.ok) fail(res.message, res.hint, res.code || 1, opt.json);
  if (opt.json) console.log(JSON.stringify(res, null, 2));
  else humanDemo(res);
  process.exit(0);
}

if (!opt.action) {
  fail('缺少子命令', `可用子命令：${Object.keys(ACTIONS).join(' / ')}\n用 ${CMD} --help 查看详情。`, 2, opt.json);
}
if (!Object.hasOwn(ACTIONS, opt.action)) {
  fail(`未知子命令：${opt.action}`, `可用子命令：${Object.keys(ACTIONS).join(' / ')}`, 2, opt.json);
}
if (!opt.input) {
  fail('缺少 --input（文件或目录）', `例：${CMD} ${opt.action} --input ./文档目录`, 2, opt.json);
}
if (opt.action === 'replace') {
  if (opt.from.length === 0) {
    fail('缺少 --from（要替换的旧文本）', `例：${CMD} replace --input ./合同 --from 旧文本 --to 新文本`, 2, opt.json);
  }
  if (opt.from.length !== opt.to.length) {
    fail(
      `--from 与 --to 数量不一致（${opt.from.length} vs ${opt.to.length}）`,
      '每一组 --from 都要配一个 --to，按顺序一一对应。',
      2,
      opt.json,
    );
  }
}

const payload = { input: opt.input, out: opt.out, from: opt.from, to: opt.to };
const res = callPython(python, opt.action, payload);
if (!res.ok) fail(res.message, res.hint, res.code || 1, opt.json);

if (opt.json) console.log(JSON.stringify(res, null, 2));
else human(res);
process.exit(0);
