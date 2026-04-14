#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成国内外海水淡化设备/系统示意图
"""

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch, Circle, Rectangle
import numpy as np
import os

plt.rcParams['font.sans-serif'] = ['Arial Unicode MS', 'SimHei', 'STHeiti']
plt.rcParams['axes.unicode_minus'] = False

output_dir = "/Users/mac/.openclaw/workspace/ppt-desalination/images"
os.makedirs(output_dir, exist_ok=True)

def create_msf_plant():
    """多级闪蒸(MSF)海水淡化厂示意图"""
    fig, ax = plt.subplots(figsize=(14, 8))
    ax.set_xlim(0, 14)
    ax.set_ylim(0, 9)
    ax.axis('off')
    
    # 标题
    ax.text(7, 8.5, '多级闪蒸（MSF）海水淡化厂', fontsize=18, fontweight='bold', ha='center')
    ax.text(7, 8, 'Multistage Flash Desalination Plant', fontsize=12, ha='center', style='italic')
    
    # 海水入口
    rect = FancyBboxPatch((0.5, 5), 2, 1.5, boxstyle="round,pad=0.05", 
                          facecolor='#CCE5FF', edgecolor='#0066CC', linewidth=2)
    ax.add_patch(rect)
    ax.text(1.5, 5.75, '海水\n进水', fontsize=11, ha='center', va='center')
    
    # 闪蒸室（多个）
    colors = ['#FFE6CC', '#FFDDBB', '#FFCC99', '#FFBB77', '#FFAA55']
    stages = ['第1级', '第2级', '第3级', '第4级', '第5级']
    for i, (stage, color) in enumerate(zip(stages, colors)):
        x = 3.5 + i * 1.8
        rect = FancyBboxPatch((x, 4.5), 1.5, 2.5, boxstyle="round,pad=0.05", 
                              facecolor=color, edgecolor='#FF6600', linewidth=2)
        ax.add_patch(rect)
        ax.text(x+0.75, 5.75, stage, fontsize=9, ha='center', va='center', fontweight='bold')
        ax.text(x+0.75, 5.25, '闪蒸\n室', fontsize=8, ha='center', va='center')
    
    # 热交换器
    rect = FancyBboxPatch((0.5, 2), 12, 1.5, boxstyle="round,pad=0.05", 
                          facecolor='#E6FFE6', edgecolor='#00AA00', linewidth=2)
    ax.add_patch(rect)
    ax.text(7, 2.75, '热交换器单元 (Heat Exchanger)', fontsize=12, ha='center', va='center', fontweight='bold')
    
    # 蒸汽入口
    rect = FancyBboxPatch((1, 7.5), 2, 1, boxstyle="round,pad=0.05", 
                          facecolor='#FFDDDD', edgecolor='#CC0000', linewidth=2)
    ax.add_patch(rect)
    ax.text(2, 8, '蒸汽\n来源', fontsize=10, ha='center', va='center')
    
    # 连接箭头
    ax.annotate('', xy=(3.5, 5.75), xytext=(2.5, 5.75),
               arrowprops=dict(arrowstyle='->', color='#0066CC', lw=2))
    
    # 产水出口
    rect = FancyBboxPatch((12, 5), 1.5, 1.5, boxstyle="round,pad=0.05", 
                          facecolor='#E6F3FF', edgecolor='#0066CC', linewidth=2)
    ax.add_patch(rect)
    ax.text(12.75, 5.75, '淡水\n产水', fontsize=10, ha='center', va='center')
    
    # 盐水出口
    rect = FancyBboxPatch((10, 1), 2, 1, boxstyle="round,pad=0.05", 
                          facecolor='#FFEEEE', edgecolor='#CC6666', linewidth=2)
    ax.add_patch(rect)
    ax.text(11, 1.5, '浓盐水\n排放', fontsize=10, ha='center', va='center')
    
    # 标注
    ax.text(7, 0.3, '特点：大规模海水淡化 | 能耗约8-10 kWh/m³ | 广泛应用于中东地区', 
            fontsize=11, ha='center', color='#666666')
    
    plt.tight_layout()
    plt.savefig(f'{output_dir}/msf_plant.png', dpi=150, bbox_inches='tight', facecolor='white')
    plt.close()
    print("Created: msf_plant.png")

def create_ro_plant():
    """反渗透(RO)海水淡化系统示意图"""
    fig, ax = plt.subplots(figsize=(14, 8))
    ax.set_xlim(0, 14)
    ax.set_ylim(0, 9)
    ax.axis('off')
    
    # 标题
    ax.text(7, 8.5, '反渗透（RO）海水淡化系统', fontsize=18, fontweight='bold', ha='center')
    ax.text(7, 8, 'Reverse Osmosis Desalination System', fontsize=12, ha='center', style='italic')
    
    # 预处理
    rect = FancyBboxPatch((0.5, 5), 2.5, 2, boxstyle="round,pad=0.05", 
                          facecolor='#FFF5E6', edgecolor='#CC8800', linewidth=2)
    ax.add_patch(rect)
    ax.text(1.75, 6.2, '预处理\n单元', fontsize=12, ha='center', va='center', fontweight='bold')
    ax.text(1.75, 5.5, '(过滤/消毒)', fontsize=9, ha='center', va='center', color='#666666')
    
    # 高压泵
    rect = FancyBboxPatch((3.5, 5.5), 2, 1.5, boxstyle="round,pad=0.05", 
                          facecolor='#FFE6E6', edgecolor='#CC0000', linewidth=2)
    ax.add_patch(rect)
    ax.text(4.5, 6.25, '高压泵', fontsize=12, ha='center', va='center', fontweight='bold')
    ax.text(4.5, 5.8, '(High Pressure)', fontsize=9, ha='center', va='center', color='#666666')
    
    # RO膜组件
    rect = FancyBboxPatch((6, 4), 3, 4, boxstyle="round,pad=0.05", 
                          facecolor='#E6F3FF', edgecolor='#0066CC', linewidth=3)
    ax.add_patch(rect)
    ax.text(7.5, 7, 'RO膜组件', fontsize=14, ha='center', va='center', fontweight='bold', color='#0066CC')
    ax.text(7.5, 6.3, 'Semi-permeable\nMembrane', fontsize=10, ha='center', va='center', style='italic')
    # 膜示意
    for i in range(3):
        ax.plot([6.3, 8.7], [5.5-i*0.4, 5.5-i*0.4], 'b-', linewidth=2, alpha=0.5)
    
    # 淡水出口
    rect = FancyBboxPatch((9.5, 6), 2, 1.5, boxstyle="round,pad=0.05", 
                          facecolor='#E6FFE6', edgecolor='#00AA00', linewidth=2)
    ax.add_patch(rect)
    ax.text(10.5, 6.75, '淡水\n(产水)', fontsize=11, ha='center', va='center', fontweight='bold')
    
    # 浓盐水出口
    rect = FancyBboxPatch((9.5, 3.5), 2, 1.5, boxstyle="round,pad=0.05", 
                          facecolor='#FFEEEE', edgecolor='#CC6666', linewidth=2)
    ax.add_patch(rect)
    ax.text(10.5, 4.25, '浓盐水\n(排水)', fontsize=11, ha='center', va='center', fontweight='bold')
    
    # 海水入口
    rect = FancyBboxPatch((0.5, 2), 2, 1, boxstyle="round,pad=0.05", 
                          facecolor='#CCE5FF', edgecolor='#0066CC', linewidth=2)
    ax.add_patch(rect)
    ax.text(1.5, 2.5, '海水进水', fontsize=10, ha='center', va='center')
    
    # 能量回收
    rect = FancyBboxPatch((6, 1.5), 3, 1.5, boxstyle="round,pad=0.05", 
                          facecolor='#FFFFCC', edgecolor='#AAAA00', linewidth=2)
    ax.add_patch(rect)
    ax.text(7.5, 2.25, '能量回收装置', fontsize=11, ha='center', va='center', fontweight='bold')
    
    # 箭头连接
    ax.annotate('', xy=(3.5, 6.25), xytext=(1.75, 6.25),
               arrowprops=dict(arrowstyle='->', color='#333333', lw=2))
    ax.annotate('', xy=(6, 6.25), xytext=(5.5, 6.25),
               arrowprops=dict(arrowstyle='->', color='#333333', lw=2))
    ax.annotate('', xy=(9.5, 6.75), xytext=(9, 6.5),
               arrowprops=dict(arrowstyle='->', color='#333333', lw=2))
    ax.annotate('', xy=(9.5, 4.25), xytext=(9, 4.5),
               arrowprops=dict(arrowstyle='->', color='#333333', lw=2))
    
    # 标注
    ax.text(7, 0.3, '特点：模块化设计 | 能耗约3-5 kWh/m³ | 全球最广泛使用的海水淡化技术', 
            fontsize=11, ha='center', color='#666666')
    
    plt.tight_layout()
    plt.savefig(f'{output_dir}/ro_plant.png', dpi=150, bbox_inches='tight', facecolor='white')
    plt.close()
    print("Created: ro_plant.png")

def create_solar_desalination():
    """太阳能海水淡化系统示意图"""
    fig, ax = plt.subplots(figsize=(14, 8))
    ax.set_xlim(0, 14)
    ax.set_ylim(0, 9)
    ax.axis('off')
    
    # 标题
    ax.text(7, 8.5, '太阳能海水淡化系统', fontsize=18, fontweight='bold', ha='center')
    ax.text(7, 8, 'Solar Desalination System', fontsize=12, ha='center', style='italic')
    
    # 太阳能板
    for i in range(3):
        for j in range(2):
            x = 1.5 + i * 1.2
            y = 6.5 + j * 0.8
            rect = patches.Rectangle((x, y), 1, 0.6, facecolor='#334455', edgecolor='#222222', linewidth=1)
            ax.add_patch(rect)
    
    ax.text(3.3, 7.9, '太阳能电池板 (Solar Panel)', fontsize=11, ha='center', va='center', fontweight='bold')
    
    # 太阳能聚光器
    ax.plot([6, 8], [7.5, 5.5], 'o-', color='#FFDD00', linewidth=4, markersize=8)
    ax.plot([6, 8], [7.2, 5.2], 'o-', color='#FFDD00', linewidth=4, markersize=8)
    ax.text(7, 6.5, '聚光器', fontsize=10, ha='center', va='center')
    
    # 蒸发器/冷凝器
    rect = FancyBboxPatch((9, 5.5), 3, 2.5, boxstyle="round,pad=0.05", 
                          facecolor='#E6FFE6', edgecolor='#00AA00', linewidth=2)
    ax.add_patch(rect)
    ax.text(10.5, 7, '太阳能蒸发-\n冷凝系统', fontsize=12, ha='center', va='center', fontweight='bold')
    ax.text(10.5, 6.2, '(Evaporation-Condensation)', fontsize=9, ha='center', va='center', style='italic')
    
    # 产水收集
    rect = FancyBboxPatch((9, 3), 3, 1.5, boxstyle="round,pad=0.05", 
                          facecolor='#E6F3FF', edgecolor='#0066CC', linewidth=2)
    ax.add_patch(rect)
    ax.text(10.5, 3.75, '淡水收集罐', fontsize=12, ha='center', va='center', fontweight='bold')
    
    # 海水箱
    rect = FancyBboxPatch((0.5, 3), 3, 2, boxstyle="round,pad=0.05", 
                          facecolor='#CCE5FF', edgecolor='#0066CC', linewidth=2)
    ax.add_patch(rect)
    ax.text(2, 4.2, '海水箱', fontsize=12, ha='center', va='center', fontweight='bold')
    ax.text(2, 3.5, '(Seawater Tank)', fontsize=9, ha='center', va='center', style='italic')
    
    # 泵
    circle = Circle((5.5, 4), 0.4, facecolor='#FFAAAA', edgecolor='#CC0000', linewidth=2)
    ax.add_patch(circle)
    ax.text(5.5, 4, '泵', fontsize=10, ha='center', va='center')
    
    # 连接箭头
    ax.annotate('', xy=(3.5, 4), xytext=(3, 4),
               arrowprops=dict(arrowstyle='->', color='#333333', lw=2))
    ax.annotate('', xy=(6, 4), xytext=(5.9, 4),
               arrowprops=dict(arrowstyle='->', color='#333333', lw=2))
    ax.annotate('', xy=(9, 4.25), xytext=(6, 4),
               arrowprops=dict(arrowstyle='->', color='#333333', lw=2))
    ax.annotate('', xy=(10.5, 5.5), xytext=(10.5, 4.5),
               arrowprops=dict(arrowstyle='->', color='#333333', lw=2))
    
    # 蒸汽标识
    ax.annotate('蒸汽', xy=(10.5, 6.8), fontsize=9, color='#666666')
    ax.annotate('', xy=(10.5, 6.5), xytext=(10.5, 6.8),
               arrowprops=dict(arrowstyle='->', color='#666666', lw=1))
    
    # 特点
    ax.text(7, 1.5, '绿色环保 | 零碳排放 | 适用于偏远地区和海岛', 
            fontsize=12, ha='center', color='#006600', fontweight='bold')
    ax.text(7, 1, 'Rising Star: 界面太阳能蒸发技术 (Interface Solar Evaporation)', 
            fontsize=10, ha='center', color='#0066CC', style='italic')
    
    # 本项目标注
    rect = FancyBboxPatch((3, 1.8), 8, 0.8, boxstyle="round,pad=0.05", 
                          facecolor='#FFE6CC', edgecolor='#FF6600', linewidth=2)
    ax.add_patch(rect)
    ax.text(7, 2.2, '本项目: 多孔钛三维界面太阳能蒸发 — 更高效！', 
            fontsize=11, ha='center', va='center', fontweight='bold', color='#FF6600')
    
    plt.tight_layout()
    plt.savefig(f'{output_dir}/solar_desalination.png', dpi=150, bbox_inches='tight', facecolor='white')
    plt.close()
    print("Created: solar_desalination.png")

def create_china_desalination():
    """中国海水淡化应用场景"""
    fig, axes = plt.subplots(1, 2, figsize=(14, 6))
    
    # 左图：沿海分布
    ax1 = axes[0]
    ax1.set_xlim(0, 10)
    ax1.set_ylim(0, 8)
    ax1.set_title('中国海水淡化厂分布', fontsize=14, fontweight='bold')
    ax1.axis('off')
    
    # 简化的中国海岸线
    coast = [(1, 2), (2, 2.5), (3, 2), (4, 2.2), (5, 2), (6, 2.5), (7, 3), (8, 3.5), (9, 4)]
    coast_x = [c[0] for c in coast]
    coast_y = [c[1] for c in coast]
    ax1.plot(coast_x, coast_y, 'b-', linewidth=2)
    ax1.fill_between(coast_x, coast_y, [0]*len(coast_y), alpha=0.3, color='#CCE5FF')
    ax1.fill_between(coast_x, coast_y, [8]*len(coast_y), alpha=0.3, color='#90EE90')
    
    # 淡化厂标注
    locations = [
        (2, 3.5, '大连\n(60万m³/d)'),
        (4, 4, '青岛\n(40万m³/d)'),
        (6, 3.5, '天津\n(50万m³/d)'),
        (8, 5, '舟山\n(30万m³/d)'),
        (7, 4.5, '宁波\n(35万m³/d)'),
        (5, 4.5, '上海\n(60万m³/d)'),
    ]
    for x, y, name in locations:
        ax1.plot(x, y, 'ro', markersize=10)
        ax1.annotate(name, xy=(x, y), xytext=(x, y+0.5), fontsize=8, ha='center', va='bottom')
    
    ax1.text(5, 0.5, '中国海水淡化产能: 约200万吨/天 | 全球第三', fontsize=10, ha='center')
    
    # 右图：应用领域
    ax2 = axes[1]
    ax2.set_xlim(0, 10)
    ax2.set_ylim(0, 8)
    ax2.set_title('海水淡化应用领域', fontsize=14, fontweight='bold')
    ax2.axis('off')
    
    sectors = [
        ('居民用水', 40, '#4477AA'),
        ('工业用水', 30, '#228833'),
        ('农业灌溉', 15, '#DDAA33'),
        ('其他', 15, '#AA4477'),
    ]
    
    colors = [s[2] for s in sectors]
    sizes = [s[1] for s in sectors]
    labels = [f"{s[0]}\n{s[1]}%" for s in sectors]
    
    wedges, texts, autotexts = ax2.pie(sizes, colors=colors, labels=labels, 
                                        autopct='', startangle=90)
    for text in texts:
        text.set_fontsize(11)
        text.set_fontweight('bold')
    
    ax2.text(5, 0.5, '主要用于沿海城市和工业基地', fontsize=10, ha='center')
    
    plt.tight_layout()
    plt.savefig(f'{output_dir}/china_desalination.png', dpi=150, bbox_inches='tight', facecolor='white')
    plt.close()
    print("Created: china_desalination.png")

def create_world_desalination():
    """全球海水淡化概况"""
    fig, ax = plt.subplots(figsize=(14, 8))
    ax.set_xlim(0, 14)
    ax.set_ylim(0, 9)
    ax.axis('off')
    
    # 标题
    ax.text(7, 8.5, '全球海水淡化产业概况', fontsize=18, fontweight='bold', ha='center')
    
    # 左侧：产能分布
    rect = FancyBboxPatch((0.5, 2), 5.5, 5.5, boxstyle="round,pad=0.1", 
                          facecolor='#F5F5F5', edgecolor='#666666', linewidth=2)
    ax.add_patch(rect)
    ax.text(3.25, 7.2, '全球海水淡化产能分布', fontsize=13, ha='center', fontweight='bold')
    
    regions = [
        ('中东和北非', 45, '#FFDD77'),
        ('北美', 20, '#77DD77'),
        ('欧洲', 15, '#77AAFF'),
        ('亚洲', 15, '#FF77AA'),
        ('其他', 5, '#AAAAAA'),
    ]
    
    y = 6.5
    for region, pct, color in regions:
        rect = FancyBboxPatch((0.8, y-0.3), pct*0.08, 0.5, boxstyle="round,pad=0.02", 
                              facecolor=color, edgecolor='#666666', linewidth=1)
        ax.add_patch(rect)
        ax.text(0.8 + pct*0.08 + 0.2, y-0.05, f'{region} ({pct}%)', fontsize=10, va='center')
        y -= 0.8
    
    # 右侧：代表国家
    rect = FancyBboxPatch((6.5, 2), 7, 5.5, boxstyle="round,pad=0.1", 
                          facecolor='#F5F5F5', edgecolor='#666666', linewidth=2)
    ax.add_patch(rect)
    ax.text(10, 7.2, '主要国家淡化产能', fontsize=13, ha='center', fontweight='bold')
    
    countries = [
        ('沙特阿拉伯', '约500万m³/d', 1),
        ('阿联酋', '约300万m³/d', 2),
        ('美国', '约280万m³/d', 3),
        ('中国', '约200万m³/d', 4),
        ('以色列', '约150万m³/d', 5),
        ('西班牙', '约100万m³/d', 6),
    ]
    
    y = 6.3
    for country, capacity, rank in countries:
        ax.text(6.8, y, f'{rank}. {country}', fontsize=10, va='center', fontweight='bold')
        ax.text(6.8, y-0.35, f'   {capacity}', fontsize=9, va='center', color='#666666')
        y -= 0.85
    
    # 底部统计
    ax.text(7, 1.2, '全球淡化水产量: 超过1亿立方米/天 | 服务人口: 超过3亿人', 
            fontsize=12, ha='center', fontweight='bold', color='#0066CC')
    ax.text(7, 0.6, '主要技术: RO (反渗透) 占比约65%, MSF (多级闪蒸) 约25%, 其他10%', 
            fontsize=10, ha='center', color='#666666')
    
    plt.tight_layout()
    plt.savefig(f'{output_dir}/world_desalination.png', dpi=150, bbox_inches='tight', facecolor='white')
    plt.close()
    print("Created: world_desalination.png")

def create_porous_titanium_design():
    """多孔钛材料设计与本项目创新点"""
    fig, axes = plt.subplots(1, 2, figsize=(14, 6))
    
    # 左图：材料设计
    ax1 = axes[0]
    ax1.set_xlim(0, 10)
    ax1.set_ylim(0, 8)
    ax1.set_title('多孔钛材料设计', fontsize=14, fontweight='bold', color='#FF6600')
    ax1.axis('off')
    
    # 微观结构
    np.random.seed(42)
    for _ in range(40):
        x = np.random.uniform(1, 9)
        y = np.random.uniform(1, 7)
        r = np.random.uniform(0.2, 0.6)
        circle = plt.Circle((x, y), r, facecolor='#DDDDDD', edgecolor='#888888', alpha=0.7)
        ax1.add_patch(circle)
    
    ax1.text(5, 0.5, '多孔结构: 孔隙率30-70% | 可调节孔径分布', fontsize=10, ha='center')
    
    # 右图：本项目创新
    ax2 = axes[1]
    ax2.set_xlim(0, 10)
    ax2.set_ylim(0, 8)
    ax2.set_title('本项目技术创新', fontsize=14, fontweight='bold', color='#0066CC')
    ax2.axis('off')
    
    innovations = [
        ('1', '三维界面设计', '增大蒸发面积3-5倍'),
        ('2', '毛细作用增强', '水分子快速输送'),
        ('3', '抗结垢表面', '延长寿命3-5倍'),
        ('4', '低温柔蒸发', '常温20-30°C运行'),
    ]
    
    y = 7
    for num, title, desc in innovations:
        circle = plt.Circle((0.8, y-0.2), 0.3, facecolor='#0066CC', edgecolor='#004488')
        ax2.add_patch(circle)
        ax2.text(0.8, y-0.2, num, fontsize=12, ha='center', va='center', color='white', fontweight='bold')
        ax2.text(1.5, y-0.1, title, fontsize=12, va='center', fontweight='bold')
        ax2.text(1.5, y-0.6, desc, fontsize=10, va='center', color='#666666')
        y -= 1.6
    
    ax2.text(5, 0.5, '目标: 能耗降低50%+ | 效率提升30%+', fontsize=11, ha='center', color='#006600', fontweight='bold')
    
    plt.tight_layout()
    plt.savefig(f'{output_dir}/porous_titanium_innovation.png', dpi=150, bbox_inches='tight', facecolor='white')
    plt.close()
    print("Created: porous_titanium_innovation.png")

# 执行所有图片生成
if __name__ == "__main__":
    print("正在生成海水淡化设备图片...")
    create_msf_plant()
    create_ro_plant()
    create_solar_desalination()
    create_china_desalination()
    create_world_desalination()
    create_porous_titanium_design()
    print(f"\n所有图片已保存到: {output_dir}")
