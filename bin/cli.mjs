#!/usr/bin/env node
/**
 * 深鲸文档批量处理 · 统一 CLI
 *
 * 用法：
 *   dsh-doc-batch docx <action> --input <路径> [--out <目录>] [--from 旧 --to 新]
 *   dsh-doc-batch xlsx <action> --input <路径> [--out <目录>]
 *   dsh-doc-batch pdf  <action> --input <路径> [--out <目录>] [--text 水印] …
 *
 * 设计：三种格式的内核（Python）与参数解析各自独立（见 bin/cli-docx|xlsx|pdf.mjs），
 * 这里只做**分发**，不复制业务逻辑 —— 保证原来单格式的行为一字不差。
 */
import { spawnSync } from 'node:child_process';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const BIN = dirname(fileURLToPath(import.meta.url));
const MAP = {
  docx: 'cli-docx.mjs', word: 'cli-docx.mjs',
  xlsx: 'cli-xlsx.mjs', excel: 'cli-xlsx.mjs',
  pdf: 'cli-pdf.mjs',
};
const CMD = 'dsh-doc-batch';

function help(code = 0) {
  console.log(`${CMD} —— 深鲸文档批量处理（Word / Excel / PDF）

用法：
  ${CMD} <格式> <操作> --input <文件或目录> [--out <目录>] [其它参数]
  ${CMD} <格式> --help          查看该格式支持的操作与参数

格式：docx（别名 word） / xlsx（别名 excel） / pdf

示例：
  ${CMD} docx replace --input ./合同 --from 深鲸科技 --to 深海鲸
  ${CMD} docx extract --input ./合同 --out ./文本
  ${CMD} xlsx merge --input ./报表 --out ./合并
  ${CMD} xlsx csv --input ./报表
  ${CMD} pdf merge --input ./扫描件 --out ./合并
  ${CMD} pdf split --input a.pdf --out ./拆分 --action page
  ${CMD} pdf watermark --input a.pdf --out ./水印 --text 机密

说明：三种格式的处理能力与原来的单格式插件完全一致；所有处理都在本机完成。`);
  process.exit(code);
}

const argv = process.argv.slice(2);
if (argv.length === 0 || argv[0] === '--help' || argv[0] === '-h') help(0);
if (argv[0] === '--version' || argv[0] === '-v') {
  console.log('0.1.0');
  process.exit(0);
}
const fmt = String(argv[0]).toLowerCase();
const target = MAP[fmt];
if (!target) {
  console.error(`未知格式：${argv[0]}\n可用：docx / xlsx / pdf（别名 word / excel）`);
  help(2);
}
const r = spawnSync(process.execPath, [join(BIN, target), ...argv.slice(1)], {
  stdio: 'inherit',
  env: { ...process.env, ELECTRON_RUN_AS_NODE: '1' },
});
process.exit(typeof r.status === 'number' ? r.status : 1);
