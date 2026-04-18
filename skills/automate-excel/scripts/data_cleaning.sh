#!/bin/bash
# 数据清洗与完整性校验 - 每日自动执行（增强版）
# 功能：一致性、准确性、有效性检查 + 多维度透视分析 + 可视化图表
# 适用场景：人力资源、财务、资产、项目等数据的清洗和完整性检查

SKILL_SCRIPTS="$HOME/.openclaw/workspace/skills/automate-excel/scripts"
DATA_DIR="$HOME/openclaw-data"
TODAY=$(date +%Y-%m-%d)
TIMESTAMP=$(date +%Y%m%d_%H%M%S)
REPORT_DIR="$DATA_DIR/cleaning"
PY_SCRIPT="$SKILL_SCRIPTS/data_quality_check.py"

# 创建输出目录
mkdir -p "$REPORT_DIR"

echo "=== 数据清洗与完整性校验开始 ==="
echo "时间：$(date)"
echo ""

# 定义各类型数据配置
declare -A DATA_TYPES
DATA_TYPES=(
    ["hr"]="员工信息表.xlsx:人力资源:部门,职级:薪资,工龄"
    ["finance"]="财务报表.xlsx:财务数据:部门,科目:金额"
    ["asset"]="资产清单.xlsx:资产数据:使用部门,资产类别:原值,折旧"
    ["project"]="项目进度表.xlsx:项目数据:负责人,项目状态:预算,进度"
)

process_type() {
    local type="$1"
    local config="$2"
    local file=$(echo "$config" | cut -d':' -f1)
    local name=$(echo "$config" | cut -d':' -f2)
    local dims=$(echo "$config" | cut -d':' -f3)
    local metrics=$(echo "$config" | cut -d':' -f4)
    
    local input_file=""
    for dir in "hr" "finance" "assets" "projects"; do
        if [ -f "$DATA_DIR/$dir/$file" ]; then
            input_file="$DATA_DIR/$dir/$file"
            break
        fi
    done
    
    if [ -z "$input_file" ]; then
        echo "⚠️ [$name] 文件未找到"
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
for type in hr finance asset project; do
    config="${DATA_TYPES[$type]}"
    process_type "$type" "$config"
done

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
