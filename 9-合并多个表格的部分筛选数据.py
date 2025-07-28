import os
import pandas as pd
from datetime import datetime
import gc
import warnings

# 忽略特定的警告
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

def process_csv_file(file_path, column_name, column_value):
    """
    处理单个 CSV 文件，筛选特定列的特定值，并返回筛选后的数据。
    """
    try:
        print(f"正在处理文件: {file_path}")
        # 读取文件的列名
        columns = pd.read_csv(file_path, nrows=0).columns
        print(f"文件 {file_path} 的列名: {repr(columns)}")

        # 清理列名，去掉多余空格
        columns = [col.strip().lower() for col in columns]

        # 检查列名是否存在
        if column_name.lower() not in columns:
            print(f"文件 {file_path} 中未找到列名 '{column_name}'，请检查列名是否正确。")
            return pd.DataFrame()
        else:
            print(f"文件 {file_path} 中找到列名 '{column_name}'，开始筛选数据。")

        # 初始化一个空的列表，用于存储筛选后的数据
        filtered_data = []
        # 每次读取 5000 行
        chunk_size = 5000
        for chunk in pd.read_csv(file_path, chunksize=chunk_size):
            # 确保列名存在
            if column_name.lower() in chunk.columns.str.lower():
                chunk = chunk.rename(columns={col: col.lower() for col in chunk.columns})
                # 筛选特定列的特定值
                mask = chunk[column_name.lower()] == column_value
                if mask.any():
                    filtered_data.append(chunk[mask])
            gc.collect()  # 手动触发垃圾回收

        # 如果筛选后的数据不为空，合并为一个 DataFrame
        if filtered_data:
            return pd.concat(filtered_data, ignore_index=True)
        else:
            print(f"文件 {file_path} 中未找到符合条件的数据。")
            return pd.DataFrame()
    except Exception as e:
        print(f"处理文件 {file_path} 时出错: {e}")
        return pd.DataFrame()

def process_csv_files(input_folder, output_folder, column_name, column_value):
    """
    处理指定文件夹中的所有 CSV 文件，筛选特定列的某个值，并将结果合并到一个 CSV 文件中。
    """
    try:
        # 获取输入文件夹中的所有文件
        all_files = [os.path.join(input_folder, f) for f in os.listdir(input_folder) if f.endswith('.csv') and not f.startswith('~$')]
        print(f"找到以下 CSV 文件: {all_files}")

        if not all_files:
            print("未找到任何 CSV 文件，请检查文件夹路径是否正确。")
            return

        # 初始化一个空的列表，用于存储所有筛选后的数据
        combined_data = []

        # 单线程处理文件
        for file_path in all_files:
            filtered_data = process_csv_file(file_path, column_name, column_value)
            if not filtered_data.empty:
                combined_data.append(filtered_data)
            gc.collect()  # 手动触发垃圾回收

        # 如果筛选后的数据不为空，合并为一个 DataFrame 并保存到输出文件
        if combined_data:
            combined_data = pd.concat(combined_data, ignore_index=True)
            
            # 自动生成输出文件名
            output_file_name = f"filtered_data_{column_value.replace('|', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            output_file_path = os.path.join(output_folder, output_file_name)
            
            combined_data.to_excel(output_file_path, index=False)
            print(f"筛选后的数据已保存到 {output_file_path}")
        else:
            print("没有找到符合条件的数据。")
    except Exception as e:
        print(f"处理文件夹 {input_folder} 时出错: {e}")

# 用户输入
def get_valid_path(prompt):
    """
    获取用户输入的有效路径。
    """
    while True:
        path = input(prompt).strip('"').strip()  # 去掉路径中的双引号和多余的空格
        if os.path.exists(path):
            return path
        else:
            print(f"路径无效: {path}。请重新输入。")

print("请输入路径时不要包含双引号。")
input_folder = get_valid_path("请输入包含 CSV 文件的文件夹路径: ").strip('"').strip()
output_folder = get_valid_path("请输入输出文件夹的路径: ").strip('"').strip()
column_name = input("请输入筛选的列名: ").strip()
column_value = input("请输入筛选的列值: ").strip()

# 调用函数
process_csv_files(input_folder, output_folder, column_name, column_value)