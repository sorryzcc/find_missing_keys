import pandas as pd

# 读取两个表格
ops_df = pd.read_excel('ops.xlsx')       # 主表
opsXY2_df = pd.read_excel('opsXY2.xlsx') # 查找表

# 打印列名，用于调试
print("ops.xlsx 列名:", ops_df.columns)
print("opsXY2.xlsx 列名:", opsXY2_df.columns)

# 确保列名一致（去除空格、统一大小写等）
ops_df.columns = ops_df.columns.str.strip()
opsXY2_df.columns = opsXY2_df.columns.str.strip()

# 如果列名是小写的 'id'，改为大写的 'ID'
if 'id' in ops_df.columns:
    ops_df.rename(columns={'id': 'ID'}, inplace=True)
if 'id' in opsXY2_df.columns:
    opsXY2_df.rename(columns={'id': 'ID'}, inplace=True)

# 使用 merge 函数进行左连接（类似于 VLOOKUP）
result_df = pd.merge(ops_df, opsXY2_df, on='ID', how='left')

# 填充缺失值（可选）
result_df['Score'] = result_df['Score'].fillna('未找到')  # 用 "未找到" 填充缺失值

# 将结果保存到新文件
result_df.to_excel('result.xlsx', index=False)

# 打印结果
print(result_df)