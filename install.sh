#!/bin/bash
# docx-editor-skill 安装脚本

set -e

echo "🚀 Installing docx-editor-skill..."

# 检测安装工具
if command -v uv &> /dev/null; then
    echo "✓ Found uv"
    INSTALLER="uv"
elif command -v pipx &> /dev/null; then
    echo "✓ Found pipx"
    INSTALLER="pipx"
else
    echo "❌ Neither uv nor pipx found. Please install one of them:"
    echo "   - uv:   pip install uv"
    echo "   - pipx: pip install pipx"
    exit 1
fi

# 获取脚本所在目录
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# 安装
echo "📦 Installing with $INSTALLER..."
if [ "$INSTALLER" = "uv" ]; then
    uv tool install --editable .
else
    pipx install --editable .
fi

# 验证
echo ""
echo "✅ Installation complete!"
echo ""
if command -v docx &> /dev/null; then
    echo "🎉 docx command is now available globally:"
    echo "   $(which docx)"
    echo ""
    echo "Try: docx --help"
else
    echo "⚠️  docx command not found in PATH."
    echo "   You may need to add ~/.local/bin to your PATH:"
    echo "   export PATH=\"\$HOME/.local/bin:\$PATH\""
fi
