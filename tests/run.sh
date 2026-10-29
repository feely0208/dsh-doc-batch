#!/bin/bash
# 深鲸文档批量处理 · 回归（统一 CLI 分发 + 三个内核仍可用）
set -uo pipefail
cd "$(dirname "$0")/.."
PASS=0; FAIL=0
ok(){ PASS=$((PASS+1)); echo "  ✅ $1"; }
bad(){ FAIL=$((FAIL+1)); echo "  ❌ $1"; }
check(){ if eval "$2" >/tmp/db.out 2>&1; then ok "$1"; else bad "$1"; sed -n '1,3p' /tmp/db.out; fi; }

echo "  深鲸文档批量处理 · 回归"

check "无参数打印帮助且退出 0" "node bin/cli.mjs >/dev/null; [ \$? -eq 0 ]"
check "帮助里列出三种格式" "node bin/cli.mjs --help | grep -q 'docx' && node bin/cli.mjs --help | grep -q 'xlsx' && node bin/cli.mjs --help | grep -q 'pdf'"
check "未知格式非 0 退出" "! node bin/cli.mjs nope --input . >/dev/null 2>&1"
check "docx 子命令帮助可用" "node bin/cli.mjs docx --help | grep -q 'Word'"
check "xlsx 子命令帮助可用" "node bin/cli.mjs xlsx --help | grep -qi 'Excel'"
check "pdf 子命令帮助可用" "node bin/cli.mjs pdf --help | grep -qi 'PDF'"
check "docx --demo 能跑通" "node bin/cli.mjs docx --demo >/dev/null"
check "xlsx --demo 能跑通" "node bin/cli.mjs xlsx --demo >/dev/null"
check "pdf --demo 能跑通" "node bin/cli.mjs pdf --demo >/dev/null"
check "别名 word/excel 可用" "node bin/cli.mjs word --help | grep -q 'Word' && node bin/cli.mjs excel --help | grep -qi 'Excel'"

echo "  ────────────────────────────────────────"
echo "  通过 $PASS · 失败 $FAIL"
[ "$FAIL" -eq 0 ]
