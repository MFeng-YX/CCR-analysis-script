import os
import pandas as pd
from datetime import datetime
import gc
import warnings
# import chardet

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

def process_csv_file(file_path, columns_to_filter, filter_values, keep_only_filtered_columns):
    try:
        print(f"正在处理文件: {file_path}")

        # 依次尝试多种编码，跳过错误行
        encodings = ['utf-8-sig', 'gbk', 'latin1']
        for enc in encodings:
            try:
                columns = pd.read_csv(file_path, nrows=0, encoding=enc).columns
                break
            except Exception:
                continue
        else:
            print(f"无法识别文件 {file_path} 的编码，跳过。")
            return pd.DataFrame()

        columns_lower = [c.strip().lower() for c in columns]
        columns_to_filter_lower = [c.strip().lower() for c in columns_to_filter]
        valid_columns = [c for c in columns_to_filter_lower if c in columns_lower]

        if not valid_columns:
            print(f"文件 {file_path} 中未找到指定的列，跳过。")
            return pd.DataFrame()

        original_col_map = {c.lower(): c for c in columns}
        filtered_chunks = []
        chunk_size = 5000

        # 使用最终确定的编码读取
        for chunk in pd.read_csv(file_path, chunksize=chunk_size, encoding=enc, on_bad_lines='skip'):
            chunk.columns = [c.strip().lower() for c in chunk.columns]
            chunk = chunk.reset_index(drop=True)
            mask = pd.Series([True] * len(chunk))

            if filter_values:
                for col, vals in filter_values.items():
                    col_lower = col.lower()
                    if col_lower in chunk.columns:
                        mask &= chunk[col_lower].astype(str).isin([str(v) for v in vals])

            filtered_chunk = chunk[mask]

            if keep_only_filtered_columns:
                filtered_chunk = filtered_chunk[[original_col_map[c] for c in valid_columns]]

            if not filtered_chunk.empty:
                filtered_chunk = filtered_chunk.rename(columns={
                    k.lower(): v for k, v in original_col_map.items()
                })
                filtered_chunks.append(filtered_chunk)

            gc.collect()

        if filtered_chunks:
            return pd.concat(filtered_chunks, ignore_index=True)
        else:
            print(f"文件 {file_path} 中未找到符合条件的数据。")
            return pd.DataFrame()
    except Exception as e:
        print(f"处理文件 {file_path} 时出错: {e}")
        return pd.DataFrame()

def process_csv_files(input_folder, output_folder, columns_to_filter, filter_values, keep_only_filtered_columns):
    try:
        all_files = [os.path.join(input_folder, f) for f in os.listdir(input_folder)
                     if f.lower().endswith('.csv') and not f.startswith('~$')]
        if not all_files:
            print("未找到任何 CSV 文件。")
            return

        combined_data = []
        for file_path in all_files:
            result = process_csv_file(file_path, columns_to_filter, filter_values, keep_only_filtered_columns)
            if not result.empty:
                combined_data.append(result)
            gc.collect()

        if combined_data:
            final_df = pd.concat(combined_data, ignore_index=True)
            suffix = "_filtered" if filter_values else "_all"
            output_file_name = f"result{suffix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            output_path = os.path.join(output_folder, output_file_name)
            final_df.to_csv(output_path, index=False)
            print(f"结果已保存到: {output_path}")
        else:
            print("没有找到任何符合条件的数据。")
    except Exception as e:
        print(f"处理文件夹时出错: {e}")

def get_valid_path(prompt):
    while True:
        path = input(prompt).strip('"').strip()
        if os.path.exists(path):
            return path
        else:
            print("路径无效，请重新输入。")

# === 用户交互 ===
print("请输入路径时不要包含双引号。")

input_folder = get_valid_path("请输入包含 CSV 文件的文件夹路径: ").strip('"').strip()
output_folder = get_valid_path("请输入输出文件夹的路径: ").strip('"').strip()

raw_columns = input("请输入要筛选的列名（多个列用英文逗号分隔）: ").strip()
columns_to_filter = [c.strip() for c in raw_columns.split(",") if c.strip()]

keep_only_filtered_columns = input("是否只保留筛选的列？输入 y 表示仅保留这些列，其他表示保留所有列: ").strip().lower() == 'y'

filter_values = {}
if input("是否根据特定值筛选？输入 y 表示是，其他表示否: ").strip().lower() == 'y':
    for col in columns_to_filter:
        val_input = input(f"请输入列 '{col}' 的筛选值（多个值用英文逗号分隔）: ").strip()
        vals = [v.strip() for v in val_input.split(',') if v.strip()]
        if vals:
            filter_values[col] = vals

# 执行
process_csv_files(input_folder, output_folder, columns_to_filter, filter_values, keep_only_filtered_columns)