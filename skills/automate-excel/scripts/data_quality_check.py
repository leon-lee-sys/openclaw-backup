#!/usr/bin/env python3
"""
数据清洗与完整性校验 - 增强版
功能：一致性、准确性、有效性检查 + 多维度透视分析 + 可视化图表
适用：人力资源、财务、资产、项目等数据
"""

import sys
import os
import argparse
from pathlib import Path
from datetime import datetime

def check_consistency(df, col_name, report):
    """一致性检查：同一实体多次出现时数据是否一致"""
    issues = []
    if col_name not in df.columns:
        return issues
    
    grouped = df.groupby(col_name)
    for val, group in grouped:
        if len(group) > 1:
            # 检查各列一致性
            for col in group.columns:
                unique_vals = group[col].dropna().unique()
                if len(unique_vals) > 1:
                    issues.append(f"  ⚠️ [{col_name}={val}] 字段[{col}]不一致: {list(unique_vals)}")
    return issues

def check_accuracy(df, rules, report):
    """准确性检查：数值范围、日期逻辑等"""
    issues = []
    for rule in rules:
        col = rule.get('column')
        rule_type = rule.get('type')
        
        if col not in df.columns:
            continue
            
        if rule_type == 'range':
            min_val = rule.get('min')
            max_val = rule.get('max')
            invalid = df[(df[col] < min_val) | (df[col] > max_val)]
            if len(invalid) > 0:
                issues.append(f"  ⚠️ [{col}] 超出范围[{min_val}-{max_val}]: {len(invalid)}条")
        
        elif rule_type == 'date_range':
            # 检查日期是否在未来或过于久远
            today = datetime.today()
            dates = pd.to_datetime(df[col], errors='coerce')
            future = dates[dates > today]
            if len(future) > 0:
                issues.append(f"  ⚠️ [{col}] 有{len(future)}条日期在未来")
    
    return issues

def check_validity(df, col_rules, report):
    """有效性检查：格式、枚举值、正则表达式"""
    issues = []
    for col, rule in col_rules.items():
        if col not in df.columns:
            continue
            
        rule_type = rule.get('type')
        
        if rule_type == 'enum':
            valid_vals = rule.get('values', [])
            invalid = df[~df[col].isin(valid_vals)]
            if len(invalid) > 0:
                issues.append(f"  ⚠️ [{col}] 无效值{len(invalid)}条，可用值: {valid_vals}")
        
        elif rule_type == 'pattern':
            import re
            pattern = rule.get('pattern')
            invalid = df[df[col].apply(lambda x: bool(x) and not re.match(pattern, str(x)))]
            if len(invalid) > 0:
                issues.append(f"  ⚠️ [{col}] 格式不符{len(invalid)}条")
        
        elif rule_type == 'not_empty':
            empty_count = df[col].isna().sum() + (df[col] == '').sum()
            if empty_count > 0:
                issues.append(f"  ⚠️ [{col}] 空值{empty_count}条")
    
    return issues

def multi_dimensional_analysis(df, dimensions, metrics, report):
    """多维度透视分析"""
    import pandas as pd
    
    results = {}
    
    # 基础统计
    results['基础统计'] = {
        '总记录数': len(df),
        '字段数': len(df.columns),
        '数据完整率': f"{(1 - df.isna().sum().sum() / (len(df) * len(df.columns))) * 100:.1f}%"
    }
    
    # 透视分析
    if dimensions and metrics:
        for metric in metrics:
            if metric not in df.columns:
                continue
            results[f'按各维度统计[{metric}]'] = {}
            for dim in dimensions:
                if dim not in df.columns:
                    continue
                grouped = df.groupby(dim)[metric].agg(['count', 'sum', 'mean', 'min', 'max'])
                results[f'按各维度统计[{metric}]'][dim] = grouped.to_dict()
    
    return results

def generate_visualization(df, output_dir, chart_types=['bar', 'line', 'pie']):
    """生成数据可视化图表"""
    import pandas as pd
    import matplotlib.pyplot as plt
    import matplotlib
    matplotlib.use('Agg')
    
    # 设置中文字体
    plt.rcParams['font.sans-serif'] = ['PingFang SC', 'Arial Unicode MS', 'SimHei']
    plt.rcParams['axes.unicode_minus'] = False
    
    charts = []
    
    # 数值字段分布图
    numeric_cols = df.select_dtypes(include=['number']).columns
    for col in numeric_cols[:5]:  # 最多5个
        try:
            fig, axes = plt.subplots(1, 2, figsize=(12, 5))
            
            # 直方图
            axes[0].hist(df[col].dropna(), bins=30, color='steelblue', alpha=0.7)
            axes[0].set_title(f'{col} 分布', fontsize=14)
            axes[0].set_xlabel(col)
            axes[0].set_ylabel('频次')
            
            # 箱线图
            axes[1].boxplot(df[col].dropna(), vert=True)
            axes[1].set_title(f'{col} 箱线图', fontsize=14)
            axes[1].set_ylabel(col)
            
            plt.tight_layout()
            path = f"{output_dir}/{col}_distribution.png"
            plt.savefig(path, dpi=150)
            plt.close()
            charts.append(path)
        except Exception as e:
            print(f"  ⚠️ 生成{col}图表失败: {e}")
    
    # 分类字段柱状图
    cat_cols = df.select_dtypes(include=['object']).columns
    for col in cat_cols[:3]:  # 最多3个
        try:
            vc = df[col].value_counts().head(10)
            if len(vc) > 0:
                fig, ax = plt.subplots(figsize=(10, 6))
                vc.plot(kind='bar', ax=ax, color='coral', alpha=0.7)
                ax.set_title(f'{col} 分布 Top10', fontsize=14)
                ax.set_xlabel(col)
                ax.set_ylabel('数量')
                plt.xticks(rotation=45, ha='right')
                plt.tight_layout()
                path = f"{output_dir}/{col}_bar.png"
                plt.savefig(path, dpi=150)
                plt.close()
                charts.append(path)
        except Exception as e:
            print(f"  ⚠️ 生成{col}图表失败: {e}")
    
    return charts

def main():
    parser = argparse.ArgumentParser(description='数据清洗与完整性校验增强版')
    parser.add_argument('--input', required=True, help='输入Excel文件路径')
    parser.add_argument('--type', required=True, choices=['hr', 'finance', 'asset', 'project'], help='数据类型')
    parser.add_argument('--output-dir', default=None, help='输出目录')
    parser.add_argument('--dimensions', default='', help='透视分析维度，逗号分隔')
    parser.add_argument('--metrics', default='', help='透视分析指标，逗号分隔')
    args = parser.parse_args()
    
    import pandas as pd
    
    input_path = args.input
    data_type = args.type
    output_dir = args.output_dir or f"~/openclaw-data/cleaning/{data_type}"
    output_dir = os.path.expanduser(output_dir)
    os.makedirs(output_dir, exist_ok=True)
    
    today = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    report_file = f"{output_dir}/report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    
    print(f"=== 数据清洗与完整性校验报告 ===")
    print(f"文件：{input_path}")
    print(f"类型：{data_type}")
    print(f"时间：{today}")
    print()
    
    # 读取数据
    try:
        df = pd.read_excel(input_path, engine='openpyxl')
    except Exception as e:
        print(f"❌ 读取文件失败: {e}")
        return
    
    # 根据数据类型配置检查规则
    col_rules = {}  # 有效性规则
    rules = []      # 准确性规则
    
    if data_type == 'hr':
        col_rules = {
            '工号': {'type': 'pattern', 'pattern': r'^[A-Z0-9]{4,12}$'},
            '部门': {'type': 'not_empty'},
            '入职日期': {'type': 'date_range'}
        }
        dimensions = ['部门', '职级']
        metrics = ['薪资', '工龄']
        
    elif data_type == 'finance':
        col_rules = {
            '凭证号': {'type': 'not_empty'},
            '金额': {'type': 'not_empty'},
            '日期': {'type': 'date_range'}
        }
        rules.append({'column': '金额', 'type': 'range', 'min': 0, 'max': 1e9})
        dimensions = ['部门', '科目']
        metrics = ['金额']
        
    elif data_type == 'asset':
        col_rules = {
            '资产编号': {'type': 'not_empty'},
            '使用部门': {'type': 'not_empty'}
        }
        dimensions = ['使用部门', '资产类别']
        metrics = ['原值', '折旧']
        
    elif data_type == 'project':
        col_rules = {
            '项目编号': {'type': 'not_empty'},
            '负责人': {'type': 'not_empty'}
        }
        dimensions = ['负责人', '项目状态']
        metrics = ['预算', '进度']
    
    # 执行检查
    with open(report_file, 'w', encoding='utf-8') as f:
        f.write(f"=== 数据清洗与完整性校验报告 ===\n")
        f.write(f"文件：{input_path}\n")
        f.write(f"类型：{data_type}\n")
        f.write(f"时间：{today}\n\n")
        
        # 1. 基础信息
        f.write("【1. 基础信息】\n")
        f.write(f"  总记录数：{len(df)}\n")
        f.write(f"  字段数：{len(df.columns)}\n")
        f.write(f"  列名：{list(df.columns)}\n\n")
        
        # 2. 一致性检查
        f.write("【2. 一致性检查】\n")
        key_col = list(col_rules.keys())[0] if col_rules else None
        if key_col:
            issues = check_consistency(df, key_col, f)
            if issues:
                for issue in issues[:10]:
                    f.write(f"{issue}\n")
            else:
                f.write("  ✅ 无一致性问题\n")
        f.write("\n")
        
        # 3. 准确性检查
        f.write("【3. 准确性检查】\n")
        if rules:
            all_issues = []
            for rule in rules:
                all_issues.extend(check_accuracy(df, [rule], f))
            if all_issues:
                for issue in all_issues[:10]:
                    f.write(f"{issue}\n")
            else:
                f.write("  ✅ 无准确性问题\n")
        else:
            f.write("  ⚠️ 未配置准确性规则\n")
        f.write("\n")
        
        # 4. 有效性检查
        f.write("【4. 有效性检查】\n")
        if col_rules:
            all_issues = check_validity(df, col_rules, f)
            if all_issues:
                for issue in all_issues[:20]:
                    f.write(f"{issue}\n")
            else:
                f.write("  ✅ 无有效性问题\n")
        else:
            f.write("  ⚠️ 未配置有效性规则\n")
        f.write("\n")
        
        # 5. 多维度透视分析
        f.write("【5. 多维度透视分析】\n")
        dims = args.dimensions.split(',') if args.dimensions else dimensions
        mets = args.metrics.split(',') if args.metrics else metrics
        if dims and mets:
            analysis = multi_dimensional_analysis(df, dims, mets, f)
            for k, v in analysis.items():
                f.write(f"  {k}：\n")
                if isinstance(v, dict):
                    for dk, dv in v.items():
                        f.write(f"    {dk}: {dv}\n")
                else:
                    f.write(f"    {v}\n")
        else:
            f.write("  ⚠️ 未配置透视维度\n")
        f.write("\n")
        
        f.write(f"报告已保存：{report_file}\n")
    
    # 生成可视化图表
    print("生成可视化图表...")
    charts = generate_visualization(df, output_dir)
    if charts:
        print(f"✅ 生成图表{len(charts)}个：")
        for c in charts:
            print(f"  - {c}")
    
    print(f"\n✅ 报告已保存：{report_file}")
    print(f"✅ 处理完成，共{len(df)}条记录")

if __name__ == '__main__':
    main()
