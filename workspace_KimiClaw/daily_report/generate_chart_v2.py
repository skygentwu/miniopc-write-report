import matplotlib.pyplot as plt
import matplotlib
import numpy as np

# 设置字体
plt.rcParams['font.family'] = 'DejaVu Sans'
plt.rcParams['axes.unicode_minus'] = False

# 蓝色科技风格配色
colors_bar = ['#1E5AA8', '#2E7AD4', '#4A90E2', '#6BA5F2', '#8FB8F7', '#B3CBFA']
colors_pie = ['#1E5AA8', '#2E7AD4', '#4A90E2', '#6BA5F2', '#8FB8F7']

# 创建图表
fig = plt.figure(figsize=(14, 6), facecolor='#F8FAFC')

# ========== 图1: 行业热度柱状图 ==========
ax1 = fig.add_subplot(121)
industries = ['Steel & Metal', 'Oil & Gas', 'Chemical', 'Mining', 'Automotive', 'New Materials']
news_counts = [12, 8, 10, 11, 9, 7]

bars = ax1.bar(industries, news_counts, color=colors_bar, edgecolor='white', linewidth=2, width=0.6)

# 添加数值标签
for bar, count in zip(bars, news_counts):
    height = bar.get_height()
    ax1.text(bar.get_x() + bar.get_width()/2., height + 0.3,
             f'{count}', ha='center', va='bottom', fontsize=12, fontweight='bold', color='#1E5AA8')

ax1.set_ylabel('News Count', fontsize=13, fontweight='bold', color='#334155')
ax1.set_title('Manufacturing AI News Distribution by Industry', fontsize=14, fontweight='bold', color='#1E3A5F', pad=15)
ax1.set_ylim(0, max(news_counts) + 3)
ax1.tick_params(axis='x', labelsize=10, rotation=15)
ax1.tick_params(axis='y', labelsize=10)
ax1.set_facecolor('#FFFFFF')
ax1.spines['top'].set_visible(False)
ax1.spines['right'].set_visible(False)
ax1.spines['left'].set_color('#CBD5E1')
ax1.spines['bottom'].set_color('#CBD5E1')
ax1.yaxis.grid(True, linestyle='--', alpha=0.3)

# ========== 图2: 技术关键词饼图 ==========
ax2 = fig.add_subplot(122)
tech_keywords = ['Predictive\nMaintenance', 'Digital Twin', 'Smart QC', 'Auto Control', 'AI LLM']
tech_counts = [28, 22, 18, 15, 17]

wedges, texts, autotexts = ax2.pie(tech_counts, labels=tech_keywords, autopct='%1.1f%%',
                                    colors=colors_pie, startangle=90,
                                    explode=(0.03, 0.03, 0.03, 0.03, 0.03),
                                    textprops={'fontsize': 10},
                                    pctdistance=0.75)

# 美化百分比文字
for autotext in autotexts:
    autotext.set_color('white')
    autotext.set_fontweight('bold')
    autotext.set_fontsize(10)

ax2.set_title('Core Technology Distribution', fontsize=14, fontweight='bold', color='#1E3A5F', pad=15)

# 添加中心圆（甜甜圈效果）
centre_circle = plt.Circle((0, 0), 0.45, fc='white')
ax2.add_patch(centre_circle)
ax2.text(0, 0, 'Tech\nMix', ha='center', va='center', fontsize=14, fontweight='bold', color='#1E5AA8')

plt.tight_layout(pad=3)
plt.savefig('/root/.openclaw/workspace/daily_report/manufacturing_ai_chart_20260307.png', 
            dpi=150, bbox_inches='tight', facecolor='#F8FAFC', edgecolor='none')
print("Chart saved to: /root/.openclaw/workspace/daily_report/manufacturing_ai_chart_20260307.png")
