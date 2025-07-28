import pandas as pd
import os
import numpy as np

PARAM_PROMPTS = {
    'table_a_path': "请输入“客户明细”表格的路径: ",
    'table_b_path': "请输入“查询结果-运单号”表格的路径: ",
    'output_dir': "请输入输出文件夹的绝对路径：",
    'date_prefix': '请输入输出文件的日期前缀：'
}

def read_file(file_path):
    file_path = file_path.strip('"')
    if file_path.endswith('.csv'):
        return pd.read_csv(file_path)
    elif file_path.endswith(('.xls', '.xlsx')):
        df = pd.read_excel(file_path)
        csv_path = file_path.rsplit('.', 1)[0] + '.csv'
        df.to_csv(csv_path, index=False)
        return pd.read_csv(csv_path)
    else:
        raise ValueError("不支持的文件格式，请提供CSV或Excel文件")

def main(table_a_path, table_b_path, output_dir, date_prefix):
    try:
        # 去除路径中的引号
        table_a_path = table_a_path.strip('"')
        table_b_path = table_b_path.strip('"')
        output_dir = output_dir.strip('"')
        
        # 检查输入路径A是否为空
        if not table_a_path or not os.path.exists(table_a_path):
            raise ValueError("客户明细表格路径无效或不存在")
        
        # 检查输入路径B是否为空
        if not table_b_path or not os.path.exists(table_b_path):
            raise ValueError("查询结果-运单号表格路径无效或不存在")
        
        # 检查输出目录是否为空
        if not output_dir:
            raise ValueError("输出目录路径不能为空")
        
        # 检查日期前缀是否为空
        if not date_prefix:
            raise ValueError("日期前缀不能为空")
        
        # 检查文件格式
        if not table_a_path.lower().endswith(('.csv', '.xlsx', '.xls')):
            raise ValueError("客户明细表格格式不支持，请使用Excel或CSV文件")
        
        if not table_b_path.lower().endswith(('.csv', '.xlsx', '.xls')):
            raise ValueError("查询结果-运单号表格格式不支持，请使用Excel或CSV文件")
    
        df_a = read_file(table_a_path)
        df_b = read_file(table_b_path)

        # 选取需要的列
        cols_a = ['省区名称','单号', '揽收网点名称', 'K码', '客户名称', '进线时间', '工单小类', '投诉/催查内容']
        df_a_selected = df_a[cols_a].copy()

        cols_b = ['运单号', '操作时间', '操作名称']
        df_b_selected = df_b[df_b['操作名称'].str.contains('入柜|入库')][cols_b]

        # 将表格A的单号与表格B的运单号进行匹配
        df_a_selected.loc[:, '进线时间'] = pd.to_datetime(df_a_selected['进线时间'])
        df_b_selected.loc[:, '操作时间'] = pd.to_datetime(df_b_selected['操作时间'])

        df_b_earliest = df_b_selected.sort_values('操作时间').drop_duplicates('运单号', keep='first')

        df_merged = pd.merge(df_a_selected, df_b_earliest, left_on='单号', right_on='运单号', how='left')
        df_merged = df_merged.rename(columns={'操作时间': '入库时间'})
        df_merged = df_merged.drop('运单号', axis=1)

        cols_order = ['省区名称', '单号', '揽收网点名称', 'K码', '客户名称', '进线时间', '入库时间', '工单小类', '投诉/催查内容']
        df_merged = df_merged[cols_order]

        # 新增几列数据
        df_merged['进线-入库时间差'] = (df_merged['进线时间'] - df_merged['入库时间']).apply(lambda x: x.total_seconds() / 3600 / 24 if not pd.isnull(x) else np.nan)

        conditions = [
            (df_merged['进线-入库时间差'] > 0),
            (df_merged['进线-入库时间差'] < 0),
            (df_merged['入库时间'].isna())
        ]
        choices = ['入库后', '入库前', '无入库']
        df_merged['入库前后'] = np.select(conditions, choices, default='')

        bins = [0, 1, 2, 3, float('inf')]
        labels = ['1天以内', '2天以内', '3天以内', '超过3天']
        df_merged['入库后进线-进线与入库时间差分布区间'] = pd.cut(
            df_merged['进线-入库时间差'],
            bins=bins,
            labels=labels,
            right=False
        )

        cols_order_new = ['省区名称', '单号', '揽收网点名称', 'K码', '客户名称', '进线时间', '入库时间', 
                        '进线-入库时间差', '入库前后', '入库后进线-进线与入库时间差分布区间', 
                        '工单小类', '投诉/催查内容']
        df_merged = df_merged[cols_order_new]

        df_merged[['入库时间', '进线-入库时间差']] = df_merged[['入库时间', '进线-入库时间差']].fillna('')

        # 确保输出文件夹存在
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        output_filename = f"{date_prefix}-客户-时间差值明细.xlsx"
        output_path = os.path.join(output_dir, output_filename)

        # 输出excel文件
        df_merged.to_excel(output_path, index=False)

        print(f"文件已成功保存到: {output_path}")

        # 显示试运行结果
        print("\n试运行结果:")
        print(df_merged.head())
        
        return {'success': True, 'message': "操作成功完成！"}
    
    except Exception as e:
        # 返回失败状态和错误信息
        return {'success': False, 'message': str(e)}

# 如果直接运行此脚本（而非被导入），则可以从命令行获取参数并调用 main 函数
if __name__ == "__main__":
    table_a_path = input("请输入“客户明细”表格的路径: ").strip('"')
    table_b_path = input("请输入“查询结果-运单号”表格的路径: ").strip('"')
    output_dir = input("请输入输出文件夹的绝对路径：").strip('"')
    date_prefix = input('请输入输出文件的日期前缀：').strip('"')
    result = main(table_a_path, table_b_path, output_dir, date_prefix)
    print(result['message'])