import json

file_path = 'e:/competition/ERP/ERP/ERP实验/市场销售报告.ipynb'
with open(file_path, 'r', encoding='utf-8') as f:
    nb = json.load(f)

merged_code = [
    "# 任务：每种商品一个图（2x6），横坐标同原图，双纵坐标分别表示单件商品净利润和总净利润\n",
    "# 净利润 = 平均价格 - 成本；总净利润 = 净利润 * 销量\n",
    "\n",
    "# 成本表读取\n",
    "cost_df = pd.read_excel('商品成本.xlsx')[['Description', 'Variable + Fixed']].copy()\n",
    "cost_df.columns = ['Material Description', 'Cost']\n",
    "cost_df['Material Description'] = cost_df['Material Description'].astype(str).str.strip()\n",
    "cost_df['Cost'] = pd.to_numeric(cost_df['Cost'], errors='coerce')\n",
    "\n",
    "# 与市场数据合并\n",
    "profit_df = grouped.copy()\n",
    "profit_df['Material Description'] = profit_df['Material Description'].astype(str).str.strip()\n",
    "profit_df = profit_df.merge(cost_df, on='Material Description', how='left')\n",
    "\n",
    "missing = profit_df.loc[profit_df['Cost'].isna(), 'Material Description'].drop_duplicates().tolist()\n",
    "if missing:\n",
    "    print('以下商品未匹配到成本，净利润将显示为NaN：', missing)\n",
    "\n",
    "profit_df['Net Profit'] = profit_df['Price'] - profit_df['Cost']\n",
    "profit_df['Total Profit'] = profit_df['Net Profit'] * profit_df['Qty']\n",
    "\n",
    "products2 = profit_df['Material Description'].dropna().unique()\n",
    "fig, axes = plt.subplots(2, 6, figsize=(36, 12))\n",
    "fig.suptitle('市场上产品净利润与总净利润变化趋势', fontsize=20, y=1.02)\n",
    "plt.subplots_adjust(hspace=0.5, wspace=0.3)\n",
    "\n",
    "# 计算全局统一的 Y 轴范围：1为单件净利润，2为总净利润\n",
    "global_min1 = profit_df['Net Profit'].min()\n",
    "global_max1 = profit_df['Net Profit'].max()\n",
    "margin1 = (global_max1 - global_min1) * 0.05\n",
    "if margin1 == 0: margin1 = 1\n",
    "\n",
    "global_min2 = profit_df['Total Profit'].min()\n",
    "global_max2 = profit_df['Total Profit'].max()\n",
    "margin2 = (global_max2 - global_min2) * 0.05\n",
    "if margin2 == 0: margin2 = 10\n",
    "\n",
    "for idx, product in enumerate(products2[:12]):\n",
    "    ax = axes[idx // 6, idx % 6]\n",
    "    d = profit_df[profit_df['Material Description'] == product].sort_values('Date')\n",
    "\n",
    "    # 左侧Y轴：单件商品净利润（紫）\n",
    "    ax.plot(d['Date'], d['Net Profit'], marker='o', linestyle='-', color='purple', label='Unit Net Profit')\n",
    "    ax.set_ylim(global_min1 - margin1, global_max1 + margin1)\n",
    "    \n",
    "    ax.set_title(product, fontsize=10)\n",
    "    ax.set_xlabel('Date')\n",
    "    ax.set_ylabel('Net Profit (Unit)', color='purple')\n",
    "    ax.tick_params(axis='y', labelcolor='purple')\n",
    "\n",
    "    ax.xaxis.set_major_formatter(mdates.DateFormatter('%m-%d'))\n",
    "    ax.xaxis.set_major_locator(mdates.DayLocator(bymonthday=[5, 10, 15, 20]))\n",
    "    plt.setp(ax.get_xticklabels(), rotation=45)\n",
    "\n",
    "    # 右侧Y轴：总净利润（绿）\n",
    "    ax2 = ax.twinx()\n",
    "    ax2.plot(d['Date'], d['Total Profit'], marker='s', linestyle='--', color='green', label='Total Net Profit')\n",
    "    ax2.set_ylabel('Total Net Profit', color='green')\n",
    "    ax2.tick_params(axis='y', labelcolor='green')\n",
    "    ax2.set_ylim(global_min2 - margin2, global_max2 + margin2)\n",
    "\n",
    "# 不足12个商品时删掉空图\n",
    "for j in range(len(products2), 12):\n",
    "    fig.delaxes(axes[j // 6, j % 6])\n",
    "\n",
    "plt.tight_layout()\n",
    "\n",
    "save_path = '商品与总净利润趋势对比图.png'\n",
    "plt.savefig(save_path, dpi=300, bbox_inches='tight') \n",
    "print(f\"图片已成功保存至: {save_path}\")\n",
    "\n",
    "plt.show()\n"
]

to_delete = []

for i, cell in enumerate(nb['cells']):
    if cell['cell_type'] == 'code':
        source_code = "".join(cell.get('source', []))
        if "商品净利润趋势图.png" in source_code:
            cell['source'] = merged_code
        elif "产品总净利润趋势图.png" in source_code:
            to_delete.append(i)

for i in reversed(to_delete):
    del nb['cells'][i]

with open(file_path, 'w', encoding='utf-8') as f:
    json.dump(nb, f, ensure_ascii=False, indent=1)

print("Update successful")
