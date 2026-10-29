/**
 * 深鲸文档批量处理 —— 宿主半边 HTTP 出口
 *
 * 一个出口管三种格式：/dsh-doc-batch/api/{ops,run}，请求里带 format。
 * 真正干活的是随包的 bin/cli.mjs（内部再分发到 cli-docx|xlsx|pdf.mjs，Python 内核不变）。
 *
 * 安全：action 白名单、参数数组传参（不拼 shell）、10 分钟超时。
 */
import { spawn } from 'node:child_process';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { existsSync, readFileSync, writeFileSync, mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';

const PKG_ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const CLI = join(PKG_ROOT, 'bin', 'cli.mjs');
const ROUTE = '/dsh-doc-batch';
/** 面板要显示版本号：直接从 package.json 读，避免客户端硬编码漂移 */
const VERSION = (() => {
  try { return JSON.parse(readFileSync(join(PKG_ROOT, 'package.json'), 'utf8')).version || ''; }
  catch { return ''; }
})();

const OPS = {
  docx: [
    { key: 'replace', label: '批量替换文本', needs: ['pairs'], hint: '每行一组「旧文本 → 新文本」，可多行' },
    { key: 'extract', label: '提取正文到 txt' },
    { key: 'merge', label: '合并成一个 docx' },
    { key: 'headers', label: '列出页眉页脚' },
    { key: 'to-md', label: '转 Markdown' },
  ],
  xlsx: [
    { key: 'merge', label: '合并工作簿（加来源列）' },
    { key: 'clean', label: '清洗：删空行 / 去重 / trim' },
    { key: 'summary', label: '列汇总报告' },
    { key: 'csv', label: '每个 sheet 导出 CSV' },
  ],
  pdf: [
    { key: 'merge', label: '合并多个 PDF' },
    { key: 'split', label: '拆分（按页 / 每 N 页）', needs: ['mode'] },
    { key: 'compress', label: '压缩' },
    { key: 'watermark', label: '加水印（文字）', needs: ['text'] },
    { key: 'extract-text', label: '提取文本' },
    { key: 'extract-images', label: '提取图片' },
  ],
};

const 上限 = 10 * 60 * 1000;

function 回(res, code, obj) {
  const body = JSON.stringify(obj);
  res.writeHead(code, { 'content-type': 'application/json; charset=utf-8', 'content-length': Buffer.byteLength(body) });
  res.end(body);
}

async function 读体(req) {
  const chunks = [];
  for await (const c of req) chunks.push(c);
  const text = Buffer.concat(chunks).toString('utf8');
  if (!text) return {};
  try { return JSON.parse(text); } catch { return {}; }
}

function 跑CLI(args) {
  return new Promise((resolve) => {
    const child = spawn(process.execPath, [CLI, ...args], {
      env: { ...process.env, ELECTRON_RUN_AS_NODE: '1' },
      windowsHide: true,
    });
    let out = '';
    const 收 = (d) => { out += String(d); if (out.length > 200_000) out = out.slice(-200_000); };
    child.stdout?.on('data', 收);
    child.stderr?.on('data', 收);
    const timer = setTimeout(() => { try { child.kill('SIGKILL'); } catch {} 收('\n[超时] 超过 10 分钟已终止。'); }, 上限);
    child.on('error', (e) => { clearTimeout(timer); resolve({ ok: false, code: -1, output: `启动失败：${e?.message || e}` }); });
    child.on('close', (code) => { clearTimeout(timer); resolve({ ok: code === 0, code: code ?? -1, output: out.trim() }); });
  });
}

/** 把面板参数翻译成 CLI 参数（按格式分别处理） */
function 组参数(b) {
  const format = String(b.format || '').toLowerCase();
  const list = OPS[format];
  if (!list) return { error: `未知格式：${b.format}` };
  const action = String(b.action || '');
  if (!list.some((o) => o.key === action)) return { error: `未知操作：${action}` };
  const input = String(b.input || '').trim();
  if (!input) return { error: '请填写要处理的文件或目录路径' };
  if (!existsSync(input)) return { error: `路径不存在：${input}` };
  const args = [format, action, '--input', input];
  const out = String(b.out || '').trim();
  if (out) args.push('--out', out);

  if (format === 'docx' && action === 'replace') {
    const 行 = String(b.pairs || '').split('\n').map((l) => l.trim()).filter(Boolean);
    if (!行.length) return { error: '批量替换至少要一组「旧 → 新」' };
    const 旧 = [], 新 = [];
    for (const l of 行) {
      const i = l.indexOf('→') >= 0 ? l.indexOf('→') : l.indexOf('->');
      if (i < 0) return { error: `这一行缺少箭头（用「旧 → 新」）：${l}` };
      const sep = l[i] === '→' ? 1 : 2;
      旧.push(l.slice(0, i).trim()); 新.push(l.slice(i + sep).trim());
    }
    for (const v of 旧) args.push('--from', v);
    for (const v of 新) args.push('--to', v);
  }
  if (format === 'pdf') {
    const n = String(b.n || '').trim();
    const text = String(b.text || '').trim();
    const opacity = String(b.opacity || '').trim();
    if (action === 'split') {
      args.push('--action', String(b.mode || 'page') === 'every' ? 'every' : 'page');
      if (n) args.push('--n', n);
    }
    if (action === 'watermark') {
      if (!text) return { error: '加水印需要填写水印文字' };
      args.push('--text', text);
      if (opacity) args.push('--opacity', opacity);
    }
    if (action === 'extract-images' && n) args.push('--dpi', n);
  }
  return { args };
}

/** 挂到宿主 webServer（ctx 需已注入 webServer） */
export function registerDocBatchApi(ctx) {
  const handler = async (req, res) => {
    try {
      const u = new URL(req.url || '/', 'http://127.0.0.1');
      const p = u.pathname;
      if (p === `${ROUTE}/api/ops` && req.method === 'GET') {
        const f = String(u.searchParams.get('format') || 'docx').toLowerCase();
        return 回(res, 200, { ok: true, format: f, version: VERSION, ops: OPS[f] || [] });
      }
      if (p === `${ROUTE}/api/upload` && req.method === 'POST') {
        const b = await 读体(req);
        const files = Array.isArray(b.files) ? b.files : [];
        if (!files.length) return 回(res, 400, { ok: false, error: '没有收到文件' });
        const dir = mkdtempSync(join(tmpdir(), 'dwh-docbatch-'));
        let total = 0; const paths = [];
        for (const f of files) {
          const name = String(f.name || 'file').replace(/[\\/:*?"<>|]/g, '_');
          const buf = Buffer.from(String(f.data || ''), 'base64');
          total += buf.length;
          if (total > 200 * 1024 * 1024) return 回(res, 413, { ok: false, error: '文件累计超过 200MB' });
          const fp = join(dir, name); writeFileSync(fp, buf); paths.push(fp);
        }
        return 回(res, 200, { ok: true, dir, path: paths.length === 1 ? paths[0] : dir, count: paths.length });
      }
      if (p === `${ROUTE}/api/run` && req.method === 'POST') {
        const b = await 读体(req);
        const { args, error } = 组参数(b);
        if (error) return 回(res, 400, { ok: false, error });
        return 回(res, 200, await 跑CLI(args));
      }
      return 回(res, 404, { ok: false, error: 'not found' });
    } catch (e) {
      return 回(res, 500, { ok: false, error: e?.message || String(e) });
    }
  };
  ctx.webServer.register({ kind: 'prefix', path: ROUTE, handler });
}
