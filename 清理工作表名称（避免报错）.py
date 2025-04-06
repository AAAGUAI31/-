

import pandas as pd
from openpyxl import load_workbook


# 读取原始Excel文件
file_path = "数据整合\鸟撞有建筑合集_20250319_110911.xlsx"
df = pd.read_excel(file_path)
# **清洗列名**（去空格、去特殊符号）
df.columns = df.columns.str.strip()

excel_file = pd.ExcelFile(file_path)
sheet_name = excel_file.sheet_names

print("工作表名称：", sheet_name)



    