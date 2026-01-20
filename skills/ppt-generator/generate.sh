#!/bin/bash
# ============================================================================
# 智能战友PPT生成器 - 快速启动脚本
# ============================================================================
#
# 使用方法:
#   ./generate.sh "报告主题" [受众] [目标]
#
# 示例:
#   ./generate.sh "2025年Q1业绩分析"
#   ./generate.sh "投资路演" investors pitch
#   ./generate.sh "产品培训" training teach
#
# ============================================================================

set -e

# 颜色定义
GREEN='\033[0;32m'
CYAN='\033[0;36m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# 进入脚本目录
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR/python"

# 加载环境变量
if [ -f "../../.env" ]; then
    export $(cat ../../.env | grep -v '^#' | xargs)
fi

# 检查参数
if [ -z "$1" ]; then
    echo -e "${CYAN}============================================${NC}"
    echo -e "${GREEN}智能战友PPT生成器${NC}"
    echo -e "${CYAN}============================================${NC}"
    echo ""
    echo "使用方法:"
    echo "  ./generate.sh \"报告主题\" [受众] [目标]"
    echo ""
    echo "受众类型:"
    echo "  c_suite   - 高管层"
    echo "  investors - 投资人"
    echo "  clients   - 客户"
    echo "  internal  - 内部团队"
    echo "  sales     - 销售场景"
    echo "  training  - 培训场景"
    echo ""
    echo "演示目标:"
    echo "  report    - 汇报"
    echo "  pitch     - 路演"
    echo "  persuade  - 说服"
    echo "  inform    - 通知"
    echo "  teach     - 培训"
    echo ""
    echo "示例:"
    echo "  ./generate.sh \"2025年Q1业绩分析\""
    echo "  ./generate.sh \"投资路演\" investors pitch"
    exit 0
fi

TOPIC="$1"
AUDIENCE="${2:-}"
GOAL="${3:-}"

echo -e "${CYAN}============================================${NC}"
echo -e "${GREEN}🧠 智能战友PPT生成器${NC}"
echo -e "${CYAN}============================================${NC}"
echo ""
echo -e "主题: ${YELLOW}$TOPIC${NC}"
[ -n "$AUDIENCE" ] && echo -e "受众: ${YELLOW}$AUDIENCE${NC}"
[ -n "$GOAL" ] && echo -e "目标: ${YELLOW}$GOAL${NC}"
echo ""

# 构建命令
CMD="python intelligent_ppt_strategist.py -t \"$TOPIC\""
[ -n "$AUDIENCE" ] && CMD="$CMD -a \"$AUDIENCE\""
[ -n "$GOAL" ] && CMD="$CMD -g \"$GOAL\""

# 执行生成
echo -e "${CYAN}开始生成...${NC}"
echo ""
eval $CMD

echo ""
echo -e "${GREEN}✅ 生成完成!${NC}"
echo -e "输出目录: ${YELLOW}./output/${NC}"
