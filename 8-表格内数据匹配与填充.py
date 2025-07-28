import pandas as pd

def process_excel():
    # 手动输入文件路径和工作表信息
    input_path = input("请输入Excel文件路径: ").strip('"')
    sheet_name = input("请输入要选取的工作表名称: ")
    a_cols = input("请输入待填充组列名，以逗号分隔: ").split(',')
    b_cols = input("请输入筛选对比组列名，以逗号分隔: ").split(',')
    output_path = input_path  # 与输入路径相同，直接覆盖

    # 去除列名两端的空白字符
    a_cols = [col.strip() for col in a_cols]
    b_cols = [col.strip() for col in b_cols]

    # 读取Excel文件
    df = pd.read_excel(input_path, sheet_name=sheet_name)

    # 筛选出A组列为空的行
    a_empty_rows = df[df[a_cols].isna().all(axis=1)]

    # 遍历每个筛选出的行
    for index, row in a_empty_rows.iterrows():
        # 获取B组列的值
        b_values = row[b_cols].values

        # 查找与B组列值完全相同的行
        matches = df[(df[b_cols] == b_values).all(axis=1)]

        # 如果找到匹配的行，填充A组列的值
        if not matches.empty:
            # 获取第一个匹配行的A组列的值
            a_values = matches.iloc[0][a_cols].values
            # 填充到当前筛选出的行
            df.loc[index, a_cols] = a_values

    # 保存修改后的数据
    with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"文件已保存到: {output_path}")

# 执行函数
process_excel()