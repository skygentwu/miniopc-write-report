import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['SimHei', 'DejaVu Sans', 'Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False

fig, ax = plt.subplots(figsize=(14, 10))

# 行业数据: (成熟度, 市场规模, 应用深度)
industries = {
    '石油化工': {'maturity': 92, 'market': 95, 'depth': 90, 'color': '#FF6B6B'},
    '钢铁冶金': {'maturity': 88, 'market': 85, 'depth': 85, 'color': '#4ECDC4'},
    '电力能源': {'maturity': 85, 'market': 90, 'depth': 88, 'color': '#45B7D1'},
    '汽车制造': {'maturity': 78, 'market': 80, 'depth': 75, 'color': '#96CEB4'},
    '矿山机械': {'maturity': 72, 'market': 70, 'depth': 70, 'color': '#FFEAA7'},
    '轨道交通': {'maturity': 75, 'market': 65, 'depth': 72, 'color': '#DDA0DD'}
}

# 绘制气泡图
for name, data in industries.items():
    ax.scatter(data['maturity'], data['market'], 
               s=data['depth']*15,  # 气泡大小表示应用深度
               c=data['color'], 
               alpha=0.7, 
               edgecolors='white', 
               linewidth=2,
               label=name)
    # 添加行业标签
    ax.annotate(name, (data['maturity'], data['market']), 
                xytext=(5, 5), textcoords='offset points',
                fontsize=11, fontweight='bold')

# 添加成熟度等级区域
ax.axvspan(80, 100, alpha=0.1, color='green', label='成熟期')
ax.axvspan(60, 80, alpha=0.1, color='orange', label='发展期')
ax.axvspan(0, 60, alpha=0.1, color='red', label='萌芽期')

# 设置坐标轴
ax.set_xlabel('技术成熟度 (%)', fontsize=14, fontweight='bold')
ax.set_ylabel('市场规模指数', fontsize=14, fontweight='bold')
ax.set_title('设备预防性维护AI智能体行业应用成熟度矩阵', fontsize=16, fontweight='bold', pad=20)

# 设置网格
ax.grid(True, alpha=0.3, linestyle='--')
ax.set_xlim(60, 100)
ax.set_ylim(55, 100)

# 添加图例说明
legend_text = "气泡大小 = 应用深度"
ax.text(0.02, 0.02, legend_text, transform=ax.transAxes, 
        fontsize=10, style='italic', bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.5))

# 添加星级标注
star_y = 98
for i, (name, data) in enumerate(industries.items()):
    stars = '★★★★★' if data['maturity'] >= 85 else '★★★★☆'
    ax.text(0.75, 0.95 - i*0.05, f"{name}: {stars}", transform=ax.transAxes,
            fontsize=10, color=data['color'], fontweight='bold')

plt.tight_layout()
plt.savefig('/root/.openclaw/workspace/chart_industry_matrix.png', dpi=150, bbox_inches='tight', 
            facecolor='white', edgecolor='none')
plt.close()

print("行业应用成熟度矩阵图已生成: chart_industry_matrix.png")
