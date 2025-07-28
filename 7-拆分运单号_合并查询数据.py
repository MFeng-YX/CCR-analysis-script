import pandas as pd
import os
from tkinter import Tk, filedialog

def split_excel_file(file_path, rows_per_file=5000):
    """
    功能1：将Excel文件按指定行数分割
    """
    try:
        # 读取Excel文件
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"读取文件出错: {e}")
        return None

    row_count = len(df)
    print(f"文件 '{file_path}' 的行数为: {row_count}")

    # 获取文件名和目录
    file_dir, file_name = os.path.split(file_path)
    file_base, file_ext = os.path.splitext(file_name)

    # 如果行数小于或等于目标行数，无需分割
    if row_count <= rows_per_file:
        print("文件行数小于等于指定行数，无需分割")
        return [file_path]

    # 分割文件
    split_files = []
    num_files = (row_count + rows_per_file - 1) // rows_per_file  # 向上取整计算文件数量

    for i in range(num_files):
        start_row = i * rows_per_file
        end_row = min((i + 1) * rows_per_file, row_count)
        subset_df = df[start_row:end_row]

        # 生成新文件名
        split_file_path = os.path.join(file_dir, f"{file_base}_{i+1}{file_ext}")
        subset_df.to_excel(split_file_path, index=False)
        split_files.append(split_file_path)
        print(f"已将第 {i+1} 部分数据保存为: {split_file_path}")

    return split_files


def merge_multiple_excel_files_by_sheet(file_paths):
    """
    功能2：合并多个Excel文件的相同工作表
    """
    # 创建一个字典来存储每个工作表名称对应的数据
    sheet_data = {}

    for file_path in file_paths:
        print(f"正在处理文件: {file_path}")
        try:
            # 获取Excel文件的所有工作表
            xls = pd.ExcelFile(file_path)
            sheet_names = xls.sheet_names
            print(f"文件 '{file_path}' 包含以下工作表: {sheet_names}")

            for sheet_name in sheet_names:
                print(f"正在读取工作表: {sheet_name}")
                try:
                    # 读取每个工作表
                    df = pd.read_excel(file_path, sheet_name=sheet_name)

                    # 如果工作表名称已存在于字典中，则追加数据
                    if sheet_name in sheet_data:
                        sheet_data[sheet_name].append(df)
                    else:
                        sheet_data[sheet_name] = [df]

                    print(f"工作表 '{sheet_name}' 读取成功，行数: {len(df)}")
                except Exception as e:
                    print(f"读取文件 '{file_path}' 的工作表 '{sheet_name}' 时出错: {e}")

        except Exception as e:
            print(f"读取文件 '{file_path}' 时出错: {e}")

    if not sheet_data:
        print("没有找到可合并的数据")
        return None

    # 合并每个工作表的数据
    merged_sheets = {}
    for sheet_name, dfs in sheet_data.items():
        merged_df = pd.concat(dfs, ignore_index=True)
        merged_sheets[sheet_name] = merged_df

    return merged_sheets


def save_merged_file_with_sheets(merged_sheets):
    """
    保存合并后的文件，包含多个工作表
    """
    # 使用tkinter的文件对话框让用户选择保存位置
    root = Tk()
    root.withdraw()
    output_path = filedialog.asksaveasfilename(
        title="选择保存合并文件的位置",
        defaultextension=".xlsx",
        filetypes=[("Excel文件", "*.xlsx")]
    )

    if not output_path:
        print("未选择保存路径，合并结果未保存")
        return None

    try:
        # 使用ExcelWriter保存多个工作表
        with pd.ExcelWriter(output_path) as writer:
            for sheet_name, merged_df in merged_sheets.items():
                merged_df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print(f"合并完成！已保存到: {output_path}")
        return output_path
    except Exception as e:
        print(f"保存合并文件时出错: {e}")
        return None


def main():
    """
    主程序：根据用户输入调用相应功能
    """
    print("请选择要执行的功能:")
    print("1. 按行数分割Excel文件")
    print("2. 合并多个Excel文件")
    
    choice = input("请输入功能编号 (1或2): ").strip()
    
    if choice == "1":
        # 功能1：按行数分割文件
        file_path = input("请输入要分割的文件路径: ").strip('"')
        split_excel_file(file_path)
    
    elif choice == "2":
        # 功能2：合并多个文件
        num_files = int(input("请输入要合并的文件数量: "))
        file_paths = []
        
        for i in range(num_files):
            file_path = input(f"请输入第 {i+1} 个文件路径: ").strip('"')
            file_paths.append(file_path)
        
        merged_sheets = merge_multiple_excel_files_by_sheet(file_paths)
        if merged_sheets is not None:
            save_merged_file_with_sheets(merged_sheets)
    
    else:
        print("无效的选择，请输入1或2")


if __name__ == "__main__":
    main()