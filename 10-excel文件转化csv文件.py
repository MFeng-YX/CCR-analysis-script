import os
import pandas as pd
import warnings


warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

def convert_excel_to_csv(input_folder, output_folder):
    """
    将指定文件夹内的所有Excel文件转换为CSV文件。

    :param input_folder: 包含Excel文件的输入文件夹路径
    :param output_folder: 输出CSV文件的文件夹路径
    """
    # 检查输入文件夹是否存在
    if not os.path.exists(input_folder):
        print(f"输入文件夹路径 {input_folder} 不存在！")
        return

    # 检查输出文件夹是否存在，如果不存在则创建
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        print(f"创建输出文件夹 {output_folder}")

    # 获取输入文件夹中的所有文件
    files = os.listdir(input_folder)
    total_files = len(files)
    processed_files = 0

    print(f"开始处理文件夹 {input_folder} 中的文件，总文件数：{total_files}")

    # 遍历输入文件夹中的所有文件
    for file_name in files:
        processed_files += 1
        print(f"正在处理文件 {processed_files}/{total_files}: {file_name}")

        # 检查文件是否为Excel文件（.xlsx 或 .xls）
        if file_name.endswith(('.xlsx', '.xls')):
            # 构建完整的文件路径
            file_path = os.path.join(input_folder, file_name)
            # 构建输出CSV文件的路径
            output_file_path = os.path.join(output_folder, file_name.replace('.xlsx', '.csv').replace('.xls', '.csv'))

            try:
                # 使用pandas读取Excel文件
                df = pd.read_excel(file_path)
                # 将数据保存为CSV文件
                df.to_csv(output_file_path, index=False)
                print(f"文件 {file_name} 已成功转换为 {output_file_path}")
            except Exception as e:
                print(f"转换文件 {file_name} 时出错：{e}")
        else:
            print(f"跳过非Excel文件：{file_name}")

    print("所有支持的Excel文件已转换完成！")

# 主程序
if __name__ == "__main__":
    # 用户输入输入文件夹路径
    input_folder = input("请输入包含Excel文件的文件夹路径：").strip('"').strip("'").strip()
    # 用户输入输出文件夹路径
    output_folder = input("请输入输出CSV文件的文件夹路径：").strip('"').strip("'").strip()

    # 调用函数进行转换
    convert_excel_to_csv(input_folder, output_folder)