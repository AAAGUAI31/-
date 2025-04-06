

import pandas as pd
import os
from datetime import datetime


# 输入和输出路径
file_path = "源数据\正式调查\鸟撞有建筑合集_2024fv2.xlsx"
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
out_path = f"数据整合/鸟撞有建筑合集_{timestamp}.xlsx"

# 需要提取的列
required_columns = [
    '编号', 'uid_p', 'uid_b', '1.团队或个人编码', '2.楼房编号', '3.是否发生鸟撞', 
    '4.鸟撞发生处周边环境', '5.鸟撞发生状况', '7.撞击面玻璃占比', '8.撞击面方向', 
    '17.完成调查日期和时间', 'userid', 'gid', '6.您所调查的建筑的编号是：', 
    '7.此建筑的名称为：', '8.此建筑的地理位置为：', '8.此建筑的地理位置为：(经度，纬度)', 
    '9.此建筑所在的地区为：:省', '9.此建筑所在的地区为：:市', '9.此建筑所在的地区为：:区', 
    '10.此建筑总共有几层楼？', '11.此建筑总体上玻璃的覆盖比例为（%）：', 
    '12.此建筑总体上有防鸟撞措施覆盖的玻璃比例为（%）：', '15.此建筑周围五米内占比最多的环境类型是', 
    '12.鸟种鉴定', '13.发生鸟撞的个体数量'
]

# 读取 Excel 文件
xls = pd.ExcelFile(file_path)
all_sheets = []


# 遍历每个工作表
for sheet_name in xls.sheet_names:
    df = pd.read_excel(xls, sheet_name=sheet_name)
    

    # 检查缺失列
    missing_cols = [col for col in required_columns if col not in df.columns]
    if missing_cols:
        print(f"⚠️ 工作表 {sheet_name} 缺少列: {missing_cols}")

    # 提取指定列
    df_filtered = df[[col for col in required_columns if col in df.columns]]

    # 增加“大地区”列标记来源
    df_filtered['大地区'] = sheet_name.strip()

    # 如果当前表有数据，加入整合列表
    if not df_filtered.empty:
        all_sheets.append(df_filtered)

# 合并所有表数据
if all_sheets:
    merged_df = pd.concat(all_sheets, ignore_index=True)
    merged_df.fillna("缺失数据", inplace=True)

    # 保存到新 Excel 文件
    merged_df.to_excel(out_path, index=False, engine='openpyxl')
    print(f"✅ 合并完成，文件已保存到: {out_path}")
else:
    print("❌ 没有找到匹配的列，未生成合并文件。")
