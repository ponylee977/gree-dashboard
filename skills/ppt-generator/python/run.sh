#!/bin/bash
# Nano Banana Pro PPT Generator - Quick Start Script

# 颜色定义
GREEN='\033[0;32m'
CYAN='\033[0;36m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

echo -e "${CYAN}"
echo "╔═══════════════════════════════════════════════════════════════╗"
echo "║          🍌 Nano Banana Pro PPT Generator                      ║"
echo "╚═══════════════════════════════════════════════════════════════╝"
echo -e "${NC}"

# 检查Python环境
if ! command -v python3 &> /dev/null; then
    echo -e "${YELLOW}警告: 未找到 python3，请先安装 Python 3.9+${NC}"
    exit 1
fi

# 切换到脚本所在目录
cd "$(dirname "$0")"

# 检查虚拟环境
if [ ! -d "venv" ]; then
    echo -e "${YELLOW}创建虚拟环境...${NC}"
    python3 -m venv venv
fi

# 激活虚拟环境
source venv/bin/activate

# 安装依赖
if [ ! -f "venv/.installed" ]; then
    echo -e "${YELLOW}安装依赖...${NC}"
    pip install -r requirements.txt
    touch venv/.installed
fi

# 运行PPT生成器
echo -e "${GREEN}启动 PPT 生成器...${NC}"
python3 nanobanana_ppt.py "$@"
