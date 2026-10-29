import { createPlugin, pkgRootFrom } from '@deepwhale-cn/dsh-guard-core/plugin';
import { registerDocBatchApi } from './api.js';

const base = createPlugin({
  name: 'doc-batch',
  pkgRoot: pkgRootFrom(import.meta.url),
  skillFile: 'skill/doc-batch.md',
  binFile: 'bin/cli.mjs',
  description: '深鲸文档批量处理：Word / Excel / PDF 一次处理整目录（替换、提取、合并、拆分、压缩、水印、导出）',
  whenToUse: '当用户要批量处理整目录的文档时使用：Word（替换/提取/合并/页眉页脚/转 Markdown）、Excel（合并/清洗/汇总/导出 CSV）、PDF（合并/拆分/压缩/水印/提取文本或图片）。典型说法：「把这批合同里的旧公司名全换掉」「这些表合并成一张」「把这些 PDF 按页拆开」。',
});

export const name = base.name;
export const inject = [...base.inject, 'webServer'];

export function apply(ctx) {
  base.apply(ctx);
  try {
    if (ctx.webServer && typeof ctx.webServer.register === 'function') {
      registerDocBatchApi(ctx);
      ctx.logger?.info?.('[doc-batch] 已挂面板出口 /dsh-doc-batch/api');
    }
  } catch (e) {
    ctx.logger?.warn?.('[doc-batch] 面板出口注册失败（CLI/skill 不受影响）: ' + (e?.message || e));
  }
}
