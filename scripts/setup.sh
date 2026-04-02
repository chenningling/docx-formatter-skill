#!/bin/bash
# setup.sh — 安装 docx-formatter 所需的 Python 依赖

set -e

echo "=== docx-formatter 环境检查 ==="

# 检查 Python
if ! command -v python3 &>/dev/null; then
    echo "✗ 未找到 python3，请先安装 Python 3.8+"
    exit 1
fi

PYTHON_VER=$(python3 --version 2>&1)
echo "✓ Python: $PYTHON_VER"

# 安装 python-docx
echo ""
echo "安装依赖：python-docx ..."
pip3 install --quiet python-docx

# 验证安装
python3 -c "from docx import Document; print('✓ python-docx 安装成功')"

echo ""
echo "=== 环境就绪，可以开始使用 ==="
