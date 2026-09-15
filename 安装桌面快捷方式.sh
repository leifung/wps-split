#!/bin/bash
# 在统信 UOS 桌面上创建“WPS 表格拆分工具”快捷方式

DIR="$(cd "$(dirname "$(readlink -f "$0" 2>/dev/null || echo "$0")")" && pwd)"

if [ -d "$HOME/Desktop" ]; then
    DESK="$HOME/Desktop"
elif [ -d "$HOME/桌面" ]; then
    DESK="$HOME/桌面"
else
    DESK="$HOME"
fi

chmod +x "$DIR/运行.sh"

cat > "$DESK/WPS表格拆分工具.desktop" <<EOF
[Desktop Entry]
Version=1.0
Type=Application
Name=WPS表格拆分工具
Name[zh_CN]=WPS表格拆分工具
Comment=按指定列的内容把一个表格拆分成多个表格
Exec=$DIR/运行.sh
Path=$DIR
Terminal=true
Categories=Office;Utility;
EOF

chmod +x "$DESK/WPS表格拆分工具.desktop" 2>/dev/null

echo "已在桌面创建快捷方式：$DESK/WPS表格拆分工具.desktop"
echo "（若图标显示为未知应用，请在其上右键 → 属性 → 勾选“允许作为程序执行”）"
read -p "按回车键退出..."
