
# --- 公司估值.ipynb ---
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

# 读取 Excel 文件
excel_file = pd.ExcelFile('公司估值.xlsx')

# 获取指定工作表中的数据
df = excel_file.parse('Sheet1')

rows, columns = df.shape

# 提取第 7 行(index=6)至第 86 行(index=85)中`Unnamed: 1`列、`Unnamed: 2`列和`Unnamed: 3`列的数据
data = df.loc[6:rows, ['Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3']]

# 设置列名
data.columns = ['Round', 'Day', 'Company Valuation']

# 转换数据类型
data['Company Valuation'] = data['Company Valuation'].str.replace(',', '').astype(float)

# 设置图片清晰度
plt.rcParams['figure.dpi'] = 300

# 设置中文字体为黑体
plt.rcParams['font.sans-serif'] = ['SimHei']
# 解决负号显示问题
plt.rcParams['axes.unicode_minus'] = False

# 设置画布大小
plt.figure(figsize=(10, 5))

# 定义颜色列表
colors = ['r', 'g', 'b', 'y']

# 遍历不同的轮次
for i, round_num in enumerate(data['Round'].unique()):
    round_data = data[data['Round'] == round_num]
    plt.plot(round_data['Day'], round_data['Company Valuation'], label=f'Round {round_num}', color=colors[i % len(colors)])

    # 每 5 天在线上标注数据
    for j in range(0, len(round_data), 5):
        plt.annotate(f'{round_data["Company Valuation"].iloc[j]:,.0f}',
                     xy=(round_data['Day'].iloc[j], round_data['Company Valuation'].iloc[j]),
                     xytext=(5, 5), textcoords='offset points', fontsize=8)

# 找出由负转正的位置并画横虚线
crossing_points = np.where(np.diff(np.sign(data['Company Valuation'])))[0] + 1
for point in crossing_points:
    plt.axhline(y=0, color='k', linestyle='--', alpha=0.5)

# 设置图表标题和坐标轴标签
plt.title('公司估值随日期变化')
plt.xlabel('日期')
plt.xticks(rotation=45)
plt.ylabel('公司估值')

# 显示图例
plt.legend()

# 保存图片
plt.savefig(f'公司估值随日期变化.png',
            bbox_inches='tight',
            dpi=300)

# 显示图表
plt.show()


# --- 市场销售报告.ipynb ---
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import seaborn as sns

# 读取Excel数据（包含Price列）
df = pd.read_excel('市场销售报告.xlsx', sheet_name='Sheet1')
rows, _ = df.shape

# 提取所需列（新增Price列）
df = df.loc[6:rows, ['Unnamed: 1', 'Unnamed: 2', 'Unnamed: 4', 'Unnamed: 6']]
df.columns = ['Date', 'Material Description', 'Qty', 'Price']

# 数据处理
df['Date'] = df['Date'].str.strip()
df['Date'] = pd.to_datetime(df['Date'], format='%m/%d')
# print(df['Date'].unique())

df['Qty'] = df['Qty'].str.replace(',', '').astype(float)
df['Price'] = df['Price'].astype(float)  # 确保价格为数值

# 按产品和日期聚合（总销量 + 平均价格）
grouped = df.groupby(['Material Description', 'Date']).agg({'Qty': 'sum', 'Price': 'mean'}).reset_index()

# 配置绘图参数
products = grouped['Material Description'].unique()
sns.set_theme()

# 设置中文字体为黑体
plt.rcParams['font.sans-serif'] = ['SimHei']
# 解决负号显示问题
plt.rcParams['axes.unicode_minus'] = False
fig, axes = plt.subplots(2, 6, figsize=(36, 12))
fig.suptitle(f'市场上产品销量与价格变化趋势', fontsize=20, y=1.02)
plt.subplots_adjust(hspace=0.5, wspace=0.3)

# 绘制双轴折线图
for idx, product in enumerate(products):
    ax = axes[idx//6, idx%6]
    data = grouped[grouped['Material Description'] == product]
    
    # 销量轴（左侧）
    ax.plot(data['Date'], data['Qty'], marker='o', linestyle='-', color='blue', label='Sales Qty')
    ax.set_title(product, fontsize=10)
    ax.set_xlabel('Date')
    ax.set_ylabel('Sales Quantity', color='blue')
    ax.tick_params(axis='y', labelcolor='blue')
    ax.set_ylim(0, 150000)
    
    # 价格轴（右侧）
    ax2 = ax.twinx()
    ax2.plot(data['Date'], data['Price'], marker='s', linestyle='--', color='red', label='Price')
    ax2.set_ylabel('Price', color='red')
    ax2.tick_params(axis='y', labelcolor='red')
    ax2.set_ylim(0, 12)
    
    # 日期格式设置
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%m-%d'))
    ax.xaxis.set_major_locator(mdates.DayLocator(bymonthday=[5, 10, 15, 20]))
    plt.setp(ax.get_xticklabels(), rotation=45)

# 显示图例
plt.legend()

# 保存图片
plt.savefig(f'市场销售报告.png',
            bbox_inches='tight',
            dpi=300)

# 调整布局
plt.tight_layout()
plt.show()


# --- 销售订单汇总.ipynb ---
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator
import seaborn as sns

# 读取数据（需要将Excel文件放在同级目录）
df = pd.read_excel("销售订单汇总.xlsx", sheet_name="Sheet1")

rows, columns = df.shape

df = df.loc[6:rows, ['Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 6']]

# 数据预处理
df.columns = ['Round', 'Day', 'Material', 'Description', 'Qty']
df['Qty'] = df['Qty'].replace(',', '', regex=True).astype(float)  # 处理千分位分隔符
df['Day'] = df['Day'].astype(int)
df['Round'] = df['Round'].astype(int)

# 获取所有产品列表（按Material排序）
products = sorted(df['Material'].unique().tolist())

# 配置绘图参数
sns.set_theme()  # 设置Seaborn的绘图风格
colors = plt.cm.tab20.colors  # 使用预定义颜色

# 按轮次绘制图表
for round_num in sorted(df['Round'].unique(), reverse=True):
    round_data = df[df['Round'] == round_num]

    if round_num == 5:  # 判断轮次是否为第 5 轮，如果是则跳过
        continue
    
    if round_data.empty:
        continue
        
    max_qty = 1
    for product in products:
        product_data = round_data[round_data['Material'] == product]
        if not product_data.empty:
            days = range(1, 21)
            qty = product_data.set_index('Day').reindex(days)['Qty'].fillna(0)
            current_max = qty.max()
            if current_max > max_qty:
                max_qty = current_max

    # 创建画布
    fig, axs = plt.subplots(2, 6, figsize=(30, 10))

    # 设置中文字体为黑体
    plt.rcParams['font.sans-serif'] = ['SimHei']
    # 解决负号显示问题
    plt.rcParams['axes.unicode_minus'] = False
    
    fig.suptitle(f'Round {round_num} 产品销量趋势', fontsize=20, y=1.02)

    # 绘制每个产品的子图
    for idx, product in enumerate(products):
        ax = axs[idx//6, idx % 6]
        product_data = round_data[round_data['Material'] == product]

        if not product_data.empty:
            # 确保日期连续
            days = range(1, 21)
            qty = product_data.set_index('Day').reindex(days)['Qty'].fillna(0)

            ax.plot(days, qty,
                    marker='o',
                    color=colors[idx],
                    linewidth=2,
                    markersize=6)

            # 设置坐标轴
            ax.xaxis.set_major_locator(MaxNLocator(integer=True))
            ax.set_xticks(range(1, 21, 2))
            ax.set_ylim(bottom=0, top=max_qty)

            # 添加标题和标签
            title = f"{product}\n{product_data['Description'].iloc[0]}"
            ax.set_title(title, fontsize=10, pad=8)
            ax.set_xlabel('Day', fontsize=8)
            ax.set_ylabel('Qty', fontsize=8)

        else:
            ax.axis('off')  # 无数据的子图留白

    # 调整布局
    plt.tight_layout()
    
    plt.subplots_adjust(hspace=0.4, wspace=0.3)

    # 保存图片
    plt.savefig(f'Round_{round_num}_Sales_Trend.png',
                bbox_inches='tight',
                dpi=300)
    plt.close()

