# 深鲸文档批量处理（@deepwhale-cn/dsh-doc-batch）

一个插件、一个面板，处理三类文档：**Word / Excel / PDF**。

## 面板

侧栏「深鲸文档批量」→ 选格式（Word / Excel / PDF）→ 选操作 → 填文件或目录 → 开始处理。
结果直接显示在面板里；所有处理都在本机完成，不上传文件。

## 安装

```bash
dsh plugin --profile web add @deepwhale-cn/dsh-doc-batch
```

也可源码安装：`dsh plugin --profile web add github:feely0208/dsh-doc-batch`。

## 命令

```bash
dsh-doc-batch docx replace --input ./合同 --from 深鲸科技 --to 深海鲸
dsh-doc-batch docx extract --input ./合同 --out ./文本
dsh-doc-batch docx merge   --input ./周报 --out ./合并
dsh-doc-batch docx headers --input ./合同
dsh-doc-batch docx to-md   --input ./合同 --out ./md

dsh-doc-batch xlsx merge   --input ./报表 --out ./合并
dsh-doc-batch xlsx clean   --input ./报表 --out ./清洗
dsh-doc-batch xlsx summary --input ./报表
dsh-doc-batch xlsx csv     --input ./报表 --out ./csv

dsh-doc-batch pdf merge          --input ./扫描件 --out ./合并
dsh-doc-batch pdf split          --input a.pdf --out ./拆分 --action page
dsh-doc-batch pdf compress       --input a.pdf --out ./压缩
dsh-doc-batch pdf watermark      --input a.pdf --out ./水印 --text 机密
dsh-doc-batch pdf extract-text   --input ./扫描件 --out ./文本
dsh-doc-batch pdf extract-images --input a.pdf --out ./图片 --dpi 150
```

格式别名：`word` = `docx`，`excel` = `xlsx`。

## 依赖

`python3` + 对应库：`python-docx`（Word）、`openpyxl`（Excel）、`pypdf`（PDF）。
缺库时 CLI 会给安装提示，面板会把提示原样显示出来。

## 说明

- 本插件由原先的三个单格式插件（docx-batch / xlsx-batch / pdf-kit）合并而来，
  **能力完全一致**，只是入口收成一个；
- 三种格式的参数解析与处理内核仍各自独立（`bin/cli-docx|xlsx|pdf.mjs` + `lib/*.py`），
  统一入口只做分发，避免"合并后行为漂移"。
