import pandas as pd
import os

PARAM_PROMPTS = {
    "input_path": "请输入文件的绝对路径：",
    "output_dir": "请输入输出文件夹的绝对路径：",
    "date_prefix": '请输入输出文件的日期前缀：'
}

def main(input_path, output_dir, date_prefix):
    try:
        # 去除路径中的引号
        input_path = input_path.strip('"')
        output_dir = output_dir.strip('"')
        
        # 检查输入路径是否为空
        if not input_path or not os.path.exists(input_path):
            raise ValueError("输入文件路径无效或不存在")
        
        # 检查输出目录是否为空
        if not output_dir:
            raise ValueError("输出目录路径不能为空")
        
        # 检查日期前缀是否为空
        if not date_prefix:
            raise ValueError("日期前缀不能为空")
        
        # 检查文件是否为 CSV 或 Excel 文件
        if not input_path.lower().endswith(('.csv', '.xlsx', '.xls')):
            raise ValueError("不支持的文件格式，请输入 CSV 或 Excel 文件")
    
        # 检查文件是否存在
        if not os.path.exists(input_path):
            print(f"文件不存在：{input_path}")
            return

        # 判断文件是否为 CSV 文件
        is_csv = input_path.lower().endswith('.csv')

        # 若不是 CSV 文件且是 Excel 文件，则转换为 CSV 文件
        if not is_csv and input_path.lower().endswith(('.xlsx', '.xls')):
            # 创建一个新的 CSV 文件路径
            csv_file_path = input_path.rsplit('.', 1)[0] + '_转换.csv'
            # 读取 Excel 文件并保存为 CSV
            df = pd.read_excel(input_path)
            df.to_csv(csv_file_path, index=False)
            print(f"已将 Excel 文件转换为 CSV 文件：{csv_file_path}")
            input_path = csv_file_path
        elif not is_csv:
            print("不支持的文件格式，请输入 CSV 或 Excel 文件")
            return

        # 读取文件
        df = pd.read_csv(input_path)

        # 打印列名
        print("列名：", df.columns)

        # 筛选需要的列
        selected_columns = ['省区名称', 'K码', '揽收网点名称', '工单来源', '工单小类', '客户名称']
        filtered_df = df[selected_columns]

        # 填充空值为"空值"，确保所有列中空值都显示为"空值"
        filtered_df = filtered_df.fillna("空值")

        # 打印筛选条件的唯一值
        print("工单来源的唯一值：", filtered_df['工单来源'].unique())
        print("工单小类的唯一值：", filtered_df['工单小类'].unique())

        # 使用.copy() 来确保操作的 DataFrame 是原始 DataFrame 或其副本
        filtered_df = filtered_df.copy()

        # 去除列值中的额外空格
        filtered_df['工单来源'] = filtered_df['工单来源'].str.strip()
        filtered_df['工单小类'] = filtered_df['工单小类'].str.strip()

        # 统一转换为小写
        filtered_df['工单来源'] = filtered_df['工单来源'].str.lower()
        filtered_df['工单小类'] = filtered_df['工单小类'].str.lower()

        # 检查满足单个条件的数据
        print("满足 '工单来源 == 客户管家小圆' 的数据行数：", filtered_df[filtered_df['工单来源'] == '客户管家小圆'].shape[0])
        print("满足 '工单小类 == 签收延误' 的数据行数：", filtered_df[filtered_df['工单小类'] == '签收延误'].shape[0])
        print("满足 '工单小类 == 派送延误' 的数据行数：", filtered_df[filtered_df['工单小类'] == '派送延误'].shape[0])

        # 筛选特定条件的行
        filtered_df = filtered_df[(filtered_df['工单来源'] == '客户管家小圆') & 
                                ((filtered_df['工单小类'] == '签收延误') | (filtered_df['工单小类'] == '派送延误'))]

        # 对筛选后的数据进行计数统计
        count_result = filtered_df.groupby(['省区名称', '揽收网点名称', 'K码', '客户名称']).size().reset_index(name='计数项：单号')

        # 按计数降序排列
        count_result = count_result.sort_values(by='计数项：单号', ascending=False)

        # 确保输出文件夹存在
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            print(f"创建输出文件夹：{output_dir}")

        # 设置输出文件名
        output_filename = f"{date_prefix}-签收派送延误筛选.xlsx"
        output_path = os.path.join(output_dir, output_filename)

        print(f"正在保存结果到文件：{output_path}")
        count_result.to_excel(output_path, index=False)

        # 打印结果
        print("筛选并计数完成！")
        print(count_result)
        print(f"结果已保存到 {output_path}")
    
        return {'success': True, 'message': "操作成功完成！"}
    
    except Exception as e:
        # 返回失败状态和错误信息
        return {'success': False, 'message': str(e)}

# 如果直接运行此脚本（而非被导入），则可以从命令行获取参数并调用 main 函数
if __name__ == "__main__":
    input_path = input("请输入文件的绝对路径：").strip('"')
    output_dir = input("请输入输出文件夹的绝对路径：").strip('"')
    date_prefix = input('请输入输出文件的日期前缀：').strip('"')
    result = main(input_path, output_dir, date_prefix)
    print(result['message'])