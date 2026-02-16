#!/bin/bash
# ============================================================
# push_to_markitdown_fork.sh
# 一键把 markitwrite 代码推送到你 fork 的 markitdown 仓库
#
# 使用方法:
#   1. 先确保你已经 fork 了 microsoft/markitdown
#   2. 在本项目根目录下运行:
#      bash push_to_markitdown_fork.sh
# ============================================================

set -e

# ---- 配置 ----
GITHUB_USER="lxistired"              # 你的 GitHub 用户名
FORK_REPO="markitdown"               # fork 出来的仓库名
BRANCH="add-markitwrite"             # 新分支名
CLONE_DIR="/tmp/markitdown-fork-$$"  # 临时克隆目录

echo "====================================="
echo "  markitwrite → markitdown fork 推送脚本"
echo "====================================="
echo ""

# Step 1: 克隆你的 fork
echo "[1/5] 克隆 ${GITHUB_USER}/${FORK_REPO} ..."
git clone --depth 1 "https://github.com/${GITHUB_USER}/${FORK_REPO}.git" "${CLONE_DIR}"
cd "${CLONE_DIR}"

# Step 2: 创建新分支
echo "[2/5] 创建分支 ${BRANCH} ..."
git checkout -b "${BRANCH}"

# Step 3: 复制 markitwrite 代码
echo "[3/5] 复制 markitwrite 代码到 packages/markitwrite/ ..."
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
mkdir -p packages/markitwrite

# 复制所有文件
cp -r "${SCRIPT_DIR}/markitwrite/pyproject.toml" packages/markitwrite/
cp -r "${SCRIPT_DIR}/markitwrite/src" packages/markitwrite/
cp -r "${SCRIPT_DIR}/markitwrite/tests" packages/markitwrite/

echo "  已复制:"
find packages/markitwrite -type f | sort | while read f; do echo "    $f"; done

# Step 4: 提交
echo "[4/5] 提交更改 ..."
git add packages/markitwrite/
git commit -m "Add markitwrite: AI virtual clipboard for pasting images into DOCX/PPTX/Markdown

Extends markitdown with reverse 'write' capability - lets AI agents insert
images into any document format with a single API call.

Architecture mirrors markitdown's converter pattern:
- DocumentWriter base class with accepts() + insert_image()
- MarkItWrite dispatcher with priority-based writer selection
- DocxWriter (python-docx), PptxWriter (python-pptx), MarkdownWriter
- CLI: markitwrite paste image.png --to output.docx
- Python API: writer.paste('screenshot.png', target='report.docx')

All 30 tests passing."

# Step 5: 推送
echo "[5/5] 推送到 GitHub ..."
git push -u origin "${BRANCH}"

echo ""
echo "====================================="
echo "  完成！"
echo "====================================="
echo ""
echo "你的 fork 已更新: https://github.com/${GITHUB_USER}/${FORK_REPO}/tree/${BRANCH}"
echo ""
echo "下一步: 你可以在 GitHub 上创建 Pull Request"
echo "  从: ${GITHUB_USER}/${FORK_REPO} (${BRANCH})"
echo "  到: microsoft/markitdown (main)"
echo ""

# 清理
cd /
rm -rf "${CLONE_DIR}"
echo "临时文件已清理。"
