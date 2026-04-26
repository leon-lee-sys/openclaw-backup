#!/bin/bash
# 数据清洗与完整性校验 - 每日自动执行（增强版）
# 功能：一致性、准确性、有效性检查 + 多维度透视分析 + 可视化图表
# 适用场景：人力资源、财务、资产、项目等数据的清洗和完整性检查

SKILL_SCRIPTS="$HOME/.openclaw/workspace/skills/automate-excel/scripts"
DATA_DIR="$HOME/Desktop/小燕子成果文件库"
TODAY=$(date +%Y-%m-%d)
TIMESTAMP=$(date +%Y%m%d_%H%M%S)
REPORT_DIR="$HOME/openclaw-data/cleaning"
PY_SCRIPT="$SKILL_SCRIPTS/data_quality_check.py"

# 创建输出目录
mkdir -p "$REPORT_DIR"

echo "=== 数据清洗与完整性校验开始 ==="
echo "时间：$(date)"
echo ""

# 定义各类型数据配置（用分号分隔的字符串）
HR_FILE="HR_数据分析看板.xlsx"
HR_NAME="人力资源"
HR_DIMS="部门,职级"
HR_METRICS="薪资,工龄"

FIN_FILE="财务_数据分析看板.xlsx"
FIN_NAME="财务数据"
FIN_DIMS="部门,科目"
FIN_METRICS="金额"

ASSET_FILE="资产_数据分析看板.xlsx"
ASSET_NAME="资产数据"
ASSET_DIMS="使用部门,资产类别"
ASSET_METRICS="原值,折旧"

PROJ_FILE="项目_数据分析看板.xlsx"
PROJ_NAME="项目数据"
PROJ_DIMS="负责人,项目状态"
PROJ_METRICS="预算,进度"

process_type() {
    local type="$1"
    local file="$2"
    local name="$3"
    local dims="$4"
    local metrics="$5"
    
    local input_file="$DATA_DIR/$file"
    
    if [ ! -f "$input_file" ]; then
        echo "⚠️ [$name] 文件未找到: $input_file"
        return
    fi
    
    echo "--- 开始处理：$name ---"
    echo "输入：$input_file"
    
    # 执行Python数据质量检查脚本
    python3 "$PY_SCRIPT" \
        --input "$input_file" \
        --type "$type" \
        --output-dir "$REPORT_DIR" \
        --dimensions "$dims" \
        --metrics "$metrics" \
        2>&1 | tee -a "$REPORT_DIR/report_${TIMESTAMP}.txt"
    
    echo ""
}

# 处理各类型数据
process_type "hr" "$HR_FILE" "$HR_NAME" "$HR_DIMS" "$HR_METRICS"
process_type "finance" "$FIN_FILE" "$FIN_NAME" "$FIN_DIMS" "$FIN_METRICS"
process_type "asset" "$ASSET_FILE" "$ASSET_NAME" "$ASSET_DIMS" "$ASSET_METRICS"
process_type "project" "$PROJ_FILE" "$PROJ_NAME" "$PROJ_DIMS" "$PROJ_METRICS"

echo ""
echo "=== 处理完成 ==="
echo "报告目录：$REPORT_DIR"
echo "今日报告：$REPORT_DIR/report_${TIMESTAMP}.txt"

# 列出生成的可视化图表
if ls "$REPORT_DIR"/*.png 1>/dev/null 2>&1; then
    echo ""
    echo "📊 生成的可视化图表："
    ls -lh "$REPORT_DIR"/*.png 2>/dev/null | awk '{print "  - "$NF" ("$5")"}'
fi
