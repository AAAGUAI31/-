import pandas as pd

# 文件路径和工作表名称
file_path = '数据整合\鸟撞有建筑合集_20250319_110911 -建筑楼层处理.xlsx'  # 替换为你的Excel文件路径
sheet_name = '10.此建筑总共有几层楼'  # 替换为你的工作表名称

# 读取Excel文件中的指定工作表
try:
    data = pd.read_excel(file_path, sheet_name=sheet_name)
except FileNotFoundError:
    print(f"文件 {file_path} 未找到，请检查路径是否正确。")
    exit()

# 假设需要统计的列名为 '楼层数'，请根据实际情况替换
column_name = '10.此建筑总共有几层楼？'

if column_name not in data.columns:
    print(f"列 {column_name} 不存在，请检查列名是否正确。")
    exit()

# 获取需要统计的列数据
floor_data = data[column_name]

# 定义区间范围
bins = [0, 5, 10, 20, 40, 60, 80, 100, float('inf')]
labels = ['1-5', '5-10', '10-20', '20-40', '40-60', '60-80', '80-100', '100+']

# 使用pandas的cut函数对数据进行分组统计
data['区间'] = pd.cut(floor_data, bins=bins, labels=labels, right=False)
count_result = data['区间'].value_counts().sort_index()

# 将统计结果转换为DataFrame
result_df = count_result.reset_index() #行添加一列索引编号 以便于我们输出
result_df.columns = ['区间', '数量']

# 将结果写入新的工作表
with pd.ExcelWriter(file_path, mode='a', engine='openpyxl') as writer:
    result_df.to_excel(writer, sheet_name='楼层统计结果', index=False)

print("统计完成，结果已写入新的工作表：统计结果")
