#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成海水淡化PPT所需的图片
"""

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch, Circle
import numpy as np
from PIL import Image, ImageDraw, ImageFont
import os

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['Arial Unicode MS', 'SimHei', 'STHeiti']
plt.rcParams['axes.unicode_minus'] = False

output_dir = "/Users/mac/.openclaw/workspace/ppt-desalination/images"
os.makedirs(output_dir, exist_ok=True)

def create_system_diagram():
    """创建系统组成示意图"""
    fig, ax = plt.subplots(1, 1, figsize=(12, 8))
    ax.set_xlim(0, 12)
    ax.set_ylim(0, 8)
    ax.axis('off')
    
    # 标题
    ax.text(6, 7.5, '多孔钛三维界面冷蒸发海水淡化系统示意图', 
            fontsize=18, fontweight='bold', ha='center', va='center')
    
    #海水箱
    rect = FancyBboxPatch((1, 3), 2.5, 2.5, boxstyle="round,pad=0.05", 
                          facecolor='#E6F3FF', edgecolor='#0066CC', linewidth=2)
    ax.add_patch(rect)
    ax.text(2.25, 4.25, '海水\n(常温)', fontsize=12, ha='center', va='center')
    
    # 多孔钛材料
    rect = FancyBboxPatch((5, 3.5), 2, 2, boxstyle="round,pad=0.05", 
                          facecolor='#FFE6CC', edgecolor='#FF6600', linewidth=3)
    ax.add_patch(rect)
    ax.text(6, 4.5, '多孔钛\n三维界面', fontsize=12, ha='center', va='center', fontweight='bold')
    ax.text(6, 3.8, '★★★', fontsize=10, ha='center', va='center', color='#FF6600')
    
    # 淡水收集
    rect = FancyBboxPatch((9, 3), 2.5, 2.5, boxstyle="round,pad=0.05", 
                          facecolor='#E6FFE6', edgecolor='#00AA00', linewidth=2)
    ax.add_patch(rect)
    ax.text(10.25, 4.25, '淡水\n收集', fontsize=12, ha='center', va='center')
    
    # 太阳能/热源
    circle = Circle((6, 6.5), 0.6, facecolor='#FFDD00', edgecolor='#FF8800', linewidth=2)
    ax.add_patch(circle)
    ax.text(6, 6.5, '☀', fontsize=20, ha='center', va='center')
    ax.text(6, 5.7, '太阳能/热源', fontsize=10, ha='center', va='center')
    
    # 箭头
    arrows = [
        (3.5, 4.5, 5, 4.5),  # 海水->多孔钛
        (7, 4.5, 9, 4.5),   # 多孔钛->淡水
        (6, 5.9, 6, 5.1),   # 热源->多孔钛
    ]
    for x1, y1, x2, y2 in arrows:
        ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                   arrowprops=dict(arrowstyle='->', color='#333333', lw=2))
    
    # 蒸汽标识
    ax.text(6.8, 5.3, '↑', fontsize=20, color='#666666')
    ax.text(6.9, 5.0, '蒸汽', fontsize=9, color='#666666')
    
    # 结垢防护标注
    ax.text(1, 2.5, '抗结垢设计', fontsize=9, color='#FF6600', style='italic')
    ax.annotate('', xy=(5, 3.8), xytext=(2.5, 2.8),
               arrowprops=dict(arrowstyle='->', color='#FF6600', lw=1.5, ls='--'))
    
    # 能耗标注
    ax.text(7.5, 6.8, '能耗降低50%+', fontsize=10, color='#0066CC', fontweight='bold')
    
    plt.tight_layout()
    plt.savefig(f'{output_dir}/system_diagram.png', dpi=150, bbox_inches='tight', 
                facecolor='white', edgecolor='none')
    plt.close()
    print(f"Created: system_diagram.png")

def create_principle_diagram():
    """创建原理图 - 三维界面蒸发机制"""
    fig, axes = plt.subplots(1, 2, figsize=(14, 6))
    
    # 左图：传统平面蒸发
    ax1 = axes[0]
    ax1.set_xlim(0, 6)
    ax1.set_ylim(0, 5)
    ax1.set_title('传统平面蒸发', fontsize=14, fontweight='bold', pad=10)
    ax1.axis('off')
    
    # 海水
    rect = FancyBboxPatch((0.5, 0.5), 5, 1.5, boxstyle="round,pad=0.05", 
                          facecolor='#CCE5FF', edgecolor='#0066CC', linewidth=2)
    ax1.add_patch(rect)
    ax1.text(3, 1.25, '海水', fontsize=11, ha='center', va='center')
    
    # 平面材料
    rect = patches.Rectangle((0.5, 2), 5, 0.3, facecolor='#CCCCCC', edgecolor='#666666', linewidth=2)
    ax1.add_patch(rect)
    ax1.text(3, 2.15, '平面蒸发层', fontsize=10, ha='center', va='center')
    
    # 热源
    ax1.text(3, 4.5, '☀ 热源', fontsize=12, ha='center', va='center')
    ax1.annotate('', xy=(3, 2.3), xytext=(3, 4),
               arrowprops=dict(arrowstyle='->', color='#FF8800', lw=2))
    
    # 蒸发效果
    ax1.text(3, 3, '↑', fontsize=16, ha='center', color='#666666')
    ax1.text(3.3, 2.7, '蒸发面积有限', fontsize=9, ha='left', color='#666666')
    
    # 右图：三维界面蒸发
    ax2 = axes[1]
    ax2.set_xlim(0, 6)
    ax2.set_ylim(0, 5)
    ax2.set_title('三维界面蒸发（本项目）', fontsize=14, fontweight='bold', pad=10, color='#0066CC')
    ax2.axis('off')
    
    # 海水
    rect = FancyBboxPatch((0.5, 0.5), 5, 1.5, boxstyle="round,pad=0.05", 
                          facecolor='#CCE5FF', edgecolor='#0066CC', linewidth=2)
    ax2.add_patch(rect)
    ax2.text(3, 1.25, '海水', fontsize=11, ha='center', va='center')
    
    # 多孔钛材料 - 3D结构
    for i in range(5):
        for j in range(3):
            x = 1 + i * 0.9
            y = 2.2 + j * 0.5
            if j == 0:
                rect = FancyBboxPatch((x, y), 0.7, 0.4, boxstyle="round,pad=0.02", 
                              facecolor='#FFE6CC', edgecolor='#FF6600', linewidth=2)
            else:
                rect = FancyBboxPatch((x, y), 0.7, 0.4, boxstyle="round,pad=0.02", 
                              facecolor='#FFDDBB', edgecolor='#FF6600', linewidth=1)
            ax2.add_patch(rect)
    
    ax2.text(3, 4, '多孔钛三维骨架', fontsize=11, ha='center', va='center', fontweight='bold', color='#FF6600')
    ax2.annotate('', xy=(3, 3.7), xytext=(3, 4.5),
               arrowprops=dict(arrowstyle='->', color='#FF6600', lw=2))
    
    # 毛细作用示意
    ax2.annotate('毛细作用\n增强', xy=(0.8, 1.8), xytext=(0.3, 1.2),
                fontsize=9, color='#0066CC',
                arrowprops=dict(arrowstyle='->', color='#0066CC', lw=1.5))
    
    # 优越性标注
    ax2.text(4.5, 3.5, '✓ 比表面积大\n✓ 热量传递快\n✓ 蒸发效率高', 
            fontsize=10, ha='left', va='center', color='#006600', linespacing=1.5)
    
    plt.tight_layout()
    plt.savefig(f'{output_dir}/principle_diagram.png', dpi=150, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close()
    print(f"Created: principle_diagram.png")

def create_tech_route():
    """创建技术路线图"""
    fig, ax = plt.subplots(figsize=(14, 6))
    ax.set_xlim(0, 14)
    ax.set_ylim(0, 6)
    ax.axis('off')
    
    # 标题
    ax.text(7, 5.5, '技术路线总览', fontsize=16, fontweight='bold', ha='center')
    
    # 阶段框
    stages = [
        ('材料设计\n第1年', 1, '#FFE6CC'),
        ('材料制备\n第1年', 3.5, '#E6F3FF'),
        ('性能测试\n第2年', 6, '#E6FFE6'),
        ('机理研究\n第2-3年', 8.5, '#FFF0FF'),
        ('模型构建\n第3年', 11, '#FFFFCC'),
        ('工程验证\n第4年', 13, '#FFE6E6'),
    ]
    
    for text, x, color in stages:
        rect = FancyBboxPatch((x-0.8, 2), 1.6, 1.8, boxstyle="round,pad=0.1", 
                              facecolor=color, edgecolor='#333333', linewidth=2)
        ax.add_patch(rect)
        ax.text(x, 2.9, text, fontsize=10, ha='center', va='center', fontweight='bold')
    
    # 连接箭头
    for i in range(len(stages)-1):
        x1 = stages[i][1] + 0.8
        x2 = stages[i+1][1] - 0.8
        y = 2.9
        ax.annotate('', xy=(x2, y), xytext=(x1, y),
                   arrowprops=dict(arrowstyle='->', color='#333333', lw=2))
    
    # 下方详细说明
    details = [
        '多孔钛材料\n设计',
        '粉末冶金/\n3D打印',
        '蒸发性能\n测试',
        '传热传质\n机理',
        '理论模型\n验证',
        '示范工程\n建设'
    ]
    
    for i, (text, x, _) in enumerate(stages):
        ax.text(x, 1.2, text, fontsize=8, ha='center', va='center', color='#666666')
    
    # 底部标注
    ax.text(7, 0.3, '材料设计 → 制备 → 测试 → 理论 → 优化 → 应用', 
            fontsize=11, ha='center', style='italic', color='#0066CC')
    
    plt.tight_layout()
    plt.savefig(f'{output_dir}/tech_route.png', dpi=150, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close()
    print(f"Created: tech_route.png")

def create_material_structure():
    """创建多孔钛材料结构示意图"""
    fig, axes = plt.subplots(1, 2, figsize=(12, 5))
    
    # 左图：微观结构示意
    ax1 = axes[0]
    ax1.set_xlim(0, 10)
    ax1.set_ylim(0, 8)
    ax1.set_title('多孔钛材料微观结构', fontsize=14, fontweight='bold')
    ax1.axis('off')
    
    # 绘制多孔结构
    np.random.seed(42)
    for _ in range(30):
        x = np.random.uniform(1, 9)
        y = np.random.uniform(1, 7)
        r = np.random.uniform(0.3, 0.8)
        circle = Circle((x, y), r, facecolor='#DDDDDD', edgecolor='#888888', linewidth=1, alpha=0.7)
        ax1.add_patch(circle)
    
    # 孔隙标注
    ax1.annotate('大孔\n(>50μm)', xy=(2, 4), fontsize=10, ha='center', color='#0066CC', fontweight='bold')
    ax1.annotate('中孔\n(10-50μm)', xy=(5, 2), fontsize=10, ha='center', color='#0066CC', fontweight='bold')
    ax1.annotate('微孔\n(<10μm)', xy=(8, 5), fontsize=10, ha='center', color='#0066CC', fontweight='bold')
    
    # 右侧说明
    ax2 = axes[1]
    ax2.set_xlim(0, 10)
    ax2.set_ylim(0, 8)
    ax2.set_title('材料特性', fontsize=14, fontweight='bold')
    ax2.axis('off')
    
    features = [
        ('🔬 高比表面积', '增大蒸发面积'),
        ('⚡ 高导热性', '快速传递热量'),
        ('🛡️ 耐腐蚀', '适应海水环境'),
        ('🎯 可调孔隙', '优化结构参数'),
    ]
    
    for i, (title, desc) in enumerate(features):
        y = 6 - i * 1.5
        ax2.text(0.5, y, title, fontsize=12, fontweight='bold', va='center')
        ax2.text(0.5, y-0.5, desc, fontsize=10, va='center', color='#666666')
    
    plt.tight_layout()
    plt.savefig(f'{output_dir}/material_structure.png', dpi=150, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close()
    print(f"Created: material_structure.png")

def create_fouling_mechanism():
    """创建结垢机理与抗垢策略图"""
    fig, axes = plt.subplots(1, 2, figsize=(14, 6))
    
    # 左图：结垢机理
    ax1 = axes[0]
    ax1.set_xlim(0, 12)
    ax1.set_ylim(0, 8)
    ax1.set_title('结垢形成机理', fontsize=14, fontweight='bold', color='#CC0000')
    ax1.axis('off')
    
    # 结垢过程
    stages = ['海水蒸发\n浓缩', '盐分析出\n成核', '晶粒长大\n沉积', '形成致密\n垢层']
    colors = ['#CCE5FF', '#FFE6CC', '#FFDDBB', '#CCCCCC']
    
    for i, (text, color) in enumerate(zip(stages, colors)):
        x = 1.5 + i * 2.8
        rect = FancyBboxPatch((x-0.8, 3), 1.6, 1.5, boxstyle="round,pad=0.05", 
                              facecolor=color, edgecolor='#666666', linewidth=2)
        ax1.add_patch(rect)
        ax1.text(x, 3.75, text, fontsize=10, ha='center', va='center', fontweight='bold')
        
        if i < len(stages) - 1:
            ax1.annotate('', xy=(x+1, 3.75), xytext=(x+0.8, 3.75),
                        arrowprops=dict(arrowstyle='->', color='#CC0000', lw=2))
    
    # 问题标注
    ax1.text(6, 1.5, '⚠️ 结垢导致：蒸发效率↓ | 孔道堵塞 | 设备损坏', 
            fontsize=11, ha='center', color='#CC0000', fontweight='bold')
    
    # 右图：抗垢策略
    ax2 = axes[1]
    ax2.set_xlim(0, 12)
    ax2.set_ylim(0, 8)
    ax2.set_title('抗结垢策略', fontsize=14, fontweight='bold', color='#006600')
    ax2.axis('off')
    
    strategies = [
        ('🧪 表面改性', '超疏水/超亲水\n图案化表面'),
        ('🎯 孔结构优化', '大孔储水\n小孔蒸发'),
        ('⚡ 外场辅助', '超声/电场\n辅助防垢'),
        ('🔄 定期维护', '自动清洗\n循环利用'),
    ]
    
    for i, (title, desc) in enumerate(strategies):
        x = 1.5 + (i % 2) * 6
        y = 5 - (i // 2) * 2.5
        
        rect = FancyBboxPatch((x-0.8, y-0.8), 2.2, 1.6, boxstyle="round,pad=0.05", 
                              facecolor='#E6FFE6', edgecolor='#006600', linewidth=2)
        ax2.add_patch(rect)
        ax2.text(x+0.3, y+0.3, title, fontsize=11, fontweight='bold', va='center')
        ax2.text(x+0.3, y-0.2, desc, fontsize=9, va='center', color='#666666')
    
    # 效果标注
    ax2.text(6, 1, '✓ 目标：结垢抑制率 ≥70% | 运行周期延长3-5倍', 
            fontsize=11, ha='center', color='#006600', fontweight='bold')
    
    plt.tight_layout()
    plt.savefig(f'{output_dir}/fouling_mechanism.png', dpi=150, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close()
    print(f"Created: fouling_mechanism.png")

def create_performance_comparison():
    """创建性能对比图"""
    fig, ax = plt.subplots(figsize=(10, 6))
    
    categories = ['蒸发速率\n(kg/m²·h)', '能量效率\n(%)', '运行周期\n(天)', '结垢抑制率\n(%)']
    
    # 传统技术数据
    traditional = [2.5, 55, 10, 25]
    # 本项目目标
    project = [5.0, 85, 30, 70]
    
    x = np.arange(len(categories))
    width = 0.35
    
    bars1 = ax.bar(x - width/2, traditional, width, label='传统技术', color='#CC6666', alpha=0.8)
    bars2 = ax.bar(x + width/2, project, width, label='本项目目标', color='#66CC66', alpha=0.8)
    
    ax.set_ylabel('数值', fontsize=12)
    ax.set_title('性能指标对比', fontsize=14, fontweight='bold')
    ax.set_xticks(x)
    ax.set_xticklabels(categories, fontsize=10)
    ax.legend(fontsize=10)
    ax.set_ylim(0, 100)
    
    # 添加数值标签
    for bar in bars1:
        height = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2., height + 2, f'{int(height)}',
               ha='center', va='bottom', fontsize=9)
    
    for bar in bars2:
        height = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2., height + 2, f'{int(height)}',
               ha='center', va='bottom', fontsize=9)
    
    plt.tight_layout()
    plt.savefig(f'{output_dir}/performance_comparison.png', dpi=150, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close()
    print(f"Created: performance_comparison.png")

def create_innovation_summary():
    """创建创新点总结图"""
    fig, ax = plt.subplots(figsize=(12, 7))
    ax.set_xlim(0, 12)
    ax.set_ylim(0, 8)
    ax.axis('off')
    
    ax.text(6, 7.5, '四大核心创新点', fontsize=18, fontweight='bold', ha='center')
    
    innovations = [
        ('🌟', '创新一', '三维界面设计', '突破传统平面蒸发瓶颈\n协同增效新概念'),
        ('🌟', '创新二', '多孔钛材料', '首创三维骨架蒸发基底\n高导热+大比表面积'),
        ('🌟', '创新三', '耦合机理', '首次系统揭示\n传热-传质-相变耦合'),
        ('🌟', '创新四', '抗垢策略', '疏水-亲水图案化\n外场辅助双重抗垢'),
    ]
    
    colors = ['#FFE6CC', '#E6F3FF', '#E6FFE6', '#FFF0FF']
    
    for i, (emoji, label, title, desc) in enumerate(innovations):
        x = 1.5 + (i % 2) * 6
        y = 5.5 - (i // 2) * 2.8
        
        rect = FancyBboxPatch((x-0.8, y-1.2), 4, 2.2, boxstyle="round,pad=0.1", 
                              facecolor=colors[i], edgecolor='#333333', linewidth=2)
        ax.add_patch(rect)
        
        ax.text(x+1.2, y+0.6, f'{emoji} {label}', fontsize=11, ha='center', fontweight='bold')
        ax.text(x+1.2, y+0.1, title, fontsize=13, ha='center', fontweight='bold', color='#0066CC')
        ax.text(x+1.2, y-0.6, desc, fontsize=9, ha='center', va='center', color='#666666', linespacing=1.3)
    
    plt.tight_layout()
    plt.savefig(f'{output_dir}/innovation_summary.png', dpi=150, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close()
    print(f"Created: innovation_summary.png")

# 执行所有图片生成
if __name__ == "__main__":
    print("正在生成PPT图片素材...")
    create_system_diagram()
    create_principle_diagram()
    create_tech_route()
    create_material_structure()
    create_fouling_mechanism()
    create_performance_comparison()
    create_innovation_summary()
    print(f"\n所有图片已保存到: {output_dir}")
