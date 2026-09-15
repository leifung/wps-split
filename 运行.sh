#!/bin/bash
# ==========================================================
#   WPS 表格按列拆分工具 —— 统信 UOS 启动脚本
#   用法一：在文件管理器里右键本文件 → 运行
#   用法二：终端执行  ./运行.sh
# ==========================================================

cd "$(dirname "$(readlink -f "$0" 2>/dev/null || echo "$0")")" || exit 1

PYTHON=""
for cmd in python3 python; do
    if command -v "$cmd" >/dev/null 2>&1; then
        PYTHON="$cmd"
        break
    fi
done

if [ -z "$PYTHON" ]; then
    echo "未找到 Python3。请先安装：sudo apt install python3"
    read -p "按回车键退出..."
    exit 1
fi

echo "=========================================="
echo "  WPS 表格按列拆分工具"
echo "  解释器：$($PYTHON -V 2>&1)"
echo "=========================================="
echo

# 优先启动图形界面；系统没有 tkinter 时自动退回命令行向导
if $PYTHON -c "import tkinter" >/dev/null 2>&1; then
    $PYTHON wps_split_gui.py
else
    echo "未检测到图形库 tkinter，已切换为命令行向导模式。"
    echo "如需图形界面，请在终端执行：sudo apt install python3-tk"
    echo
    $PYTHON wps_split.py
fi

echo
read -p "处理完毕，按回车键关闭窗口..."
