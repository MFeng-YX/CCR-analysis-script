import pandas as pd
import os

PARAM_PROMPTS = {
    'input_path': "请输入“客户-时间差值明细”文件路径：",
    'output_dir': "请输入输出文件夹的绝对路径：",
    'date_prefix': '请输入输出文件的日期前缀：'
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
        
        # 检查文件格式
        if not input_path.lower().endswith(('.csv', '.xlsx', '.xls')):
            raise ValueError("不支持的文件格式，请提供CSV或Excel文件")
    
        # 检查文件格式并读取文件
        if input_path.endswith('.csv'):
            df = pd.read_csv(input_path)
        elif input_path.endswith(('.xls', '.xlsx')):
            df = pd.read_excel(input_path)
            csv_path = input_path.replace('.xlsx', '.csv').replace('.xls', '.csv')
            df.to_csv(csv_path, index=False)
            df = pd.read_csv(csv_path)
        else:
            print("不支持的文件格式")
            return

        # 处理缺失值，指定每个列的数据类型
        for col in df.columns:
            if pd.api.types.is_numeric_dtype(df[col]):
                df.loc[:, col] = df[col].fillna(0)
            else:
                df.loc[:, col] = df[col].fillna('')

        # 创建第一个工作表的数据
        sheet1_data = df.groupby(['省区名称', '揽收网点名称', 'K码', '客户名称'])['入库前后'].value_counts().unstack(fill_value=0).reset_index()
        sheet1_data['总计'] = sheet1_data[['入库后', '入库前', '无入库']].sum(axis=1)
        sheet1_data[['入库后占比', '入库前占比', '无入库占比']] = sheet1_data[['入库后', '入库前', '无入库']].div(sheet1_data['总计'], axis=0)

        # 调整列顺序
        sheet1_data = sheet1_data[['揽收网点名称', 'K码', '客户名称', '无入库', '无入库占比', '入库前', '入库前占比', '入库后', '入库后占比', '总计']]
        sheet1_data.rename(columns={'入库后': '入库后进线', '入库前': '入库前进线', '无入库': '无入库进线'}, inplace=True)

        # 添加总计行
        total_row1 = {
            '省区名称': '',
            '揽收网点名称': '总计',
            'K码': '',
            '客户名称': '',
            '无入库进线': sheet1_data['无入库进线'].sum(),
            '无入库占比': sheet1_data['无入库进线'].sum() / sheet1_data['总计'].sum(),
            '入库前进线': sheet1_data['入库前进线'].sum(),
            '入库前占比': sheet1_data['入库前进线'].sum() / sheet1_data['总计'].sum(),
            '入库后进线': sheet1_data['入库后进线'].sum(),
            '入库后占比': sheet1_data['入库后进线'].sum() / sheet1_data['总计'].sum(),
            '总计': sheet1_data['总计'].sum()
        }
        sheet1_data = pd.concat([sheet1_data, pd.DataFrame([total_row1])], ignore_index=True)

        # 创建第二个工作表的数据
        sheet2_data = df.groupby(['省区名称', '揽收网点名称', 'K码', '客户名称'])['入库后进线-进线与入库时间差分布区间'].value_counts().unstack(fill_value=0).reset_index()
        sheet2_data['总计'] = sheet2_data[['1天以内', '2天以内', '3天以内', '超过3天']].sum(axis=1)
        sheet2_data[['1天以内占比', '2天以内占比', '3天以内占比', '超3天占比']] = sheet2_data[['1天以内', '2天以内', '3天以内', '超过3天']].div(sheet2_data['总计'], axis=0)
        sheet2_data = sheet2_data[['揽收网点名称', 'K码', '客户名称', '1天以内', '1天以内占比', '2天以内', '2天以内占比', '3天以内', '3天以内占比', '超过3天', '超3天占比', '总计']]

        # 添加总计行
        total_row2 = {
            '省区名称': '',
            '揽收网点名称': '总计',
            'K码': '',
            '客户名称': '',
            '1天以内': sheet2_data['1天以内'].sum(),
            '1天以内占比': sheet2_data['1天以内'].sum() / sheet2_data['总计'].sum(),
            '2天以内': sheet2_data['2天以内'].sum(),
            '2天以内占比': sheet2_data['2天以内'].sum() / sheet2_data['总计'].sum(),
            '3天以内': sheet2_data['3天以内'].sum(),
            '3天以内占比': sheet2_data['3天以内'].sum() / sheet2_data['总计'].sum(),
            '超过3天': sheet2_data['超过3天'].sum(),
            '超3天占比': sheet2_data['超过3天'].sum() / sheet2_data['总计'].sum(),
            '总计': sheet2_data['总计'].sum()
        }
        sheet2_data = pd.concat([sheet2_data, pd.DataFrame([total_row2])], ignore_index=True)

        # 确保输出文件夹存在
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        output_filename = f"{date_prefix}-客户-时间差值明细分析.xlsx"
        output_path = os.path.join(output_dir, output_filename)

        # 将两个工作表写入 Excel 文件
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            sheet1_data.to_excel(writer, sheet_name='入库进线分析', index=False)
            sheet2_data.to_excel(writer, sheet_name='入库后催件时间分析', index=False)

        print("分析完成，结果已保存到：", output_path)
        
        return {'success': True, 'message': "操作成功完成！"}
    
    except Exception as e:
        # 返回失败状态和错误信息
        return {'success': False, 'message': str(e)}

# 如果直接运行此脚本（而非被导入），则可以从命令行获取参数并调用 main 函数
if __name__ == "__main__":
    input_path = input("请输入“客户-时间差值明细”文件路径：").strip('"')
    output_dir = input("请输入输出文件夹的绝对路径：").strip('"')
    date_prefix = input('请输入输出文件的日期前缀：').strip('"')
    result = main(input_path, output_dir, date_prefix)
    print(result['message'])