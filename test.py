import pandas as pd

# 读取两个表格
ops_df = pd.read_excel('ops.xlsx')       # 主表
opsXY2_df = pd.read_excel('opsXY2.xlsx') # 查找表

# 打印列名，用于调试
print("ops.xlsx 列名:", ops_df.columns.tolist())
print("opsXY2.xlsx 列名:", opsXY2_df.columns.tolist())

# 使用 merge 函数进行左连接，选择需要的列（PO 和 Version）
result_df = pd.merge(
    ops_df,
    opsXY2_df[['Key', 'PO', 'Version']],  # 只选择需要的列
    on='Key',                             # 两个表格的连接键都是 'Key'
    how='left'
)

# 填充缺失值（可选）
result_df['PO'] = result_df['PO'].fillna('未找到')
result_df['Version'] = result_df['Version'].fillna('未找到')

# 将结果保存到新文件
result_df.to_excel('result.xlsx', index=False)

# 打印结果
print(result_df)