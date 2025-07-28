import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os
from datetime import datetime
import io

"""
6. 时间差值_图表分析_客户多日维度
"""
PARAM_PROMPTS = {
    'direct_table_path': "直接表格数据路径（如果没有，输入n）：",
    'output_dir': "输出文件夹路径：",
    'file_paths': "请输入多个表格文件的路径（用逗号分隔）："
}

plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False

def main(direct_table_path=None, output_dir=None, file_paths=None):
    try:
        # 去除路径中的引号
        direct_table_path = direct_table_path.strip('"') if direct_table_path else None
        output_dir = output_dir.strip('"') if output_dir else None
        
        # 检查直接表格路径
        if direct_table_path and direct_table_path.lower() != 'n':
            if not direct_table_path or not os.path.exists(direct_table_path):
                raise ValueError("直接表格数据路径无效或不存在")
            if not direct_table_path.lower().endswith('.csv'):
                raise ValueError("直接表格数据格式不支持，请提供CSV文件")
        
        # 检查输出目录是否为空
        if not output_dir:
            raise ValueError("输出目录路径不能为空")
        
        if direct_table_path and direct_table_path.lower() != 'n':
            combined_df = pd.read_csv(direct_table_path)
        else:
            if file_paths is None:
                raise ValueError("file_paths 参数不能为空")
            
            # 将文件路径列表从字符串转换为列表
            file_paths = [path.strip('"') for path in file_paths.split(',')]
            
            # 检查每个文件路径是否有效
            for path in file_paths:
                if not os.path.exists(path):
                    raise ValueError(f"文件路径无效或不存在: {path}")
                if not path.lower().endswith('.csv'):
                    raise ValueError(f"文件格式不支持，请提供CSV文件: {path}")
            
            df_list = [pd.read_csv(file_path) for file_path in file_paths]
            combined_df = pd.concat(df_list, ignore_index=True)
            combined_df.to_csv(os.path.join(output_dir, '周度汇总.csv'), index=False)

        selected_data = combined_df[['客户名称', '进线时间', '进线-入库时间差']]
        customer_counts = selected_data['客户名称'].value_counts()
        customers_with_counts_gt1 = customer_counts[customer_counts > 1].index
        filtered_data = selected_data[selected_data['客户名称'].isin(customers_with_counts_gt1)]
        filtered_data['进线时间日期'] = pd.to_datetime(filtered_data['进线时间']).dt.date

        customer_date_count = {}
        for customer in customers_with_counts_gt1:
            customer_data = filtered_data[filtered_data['客户名称'] == customer]
            date_count = {}
            for date in customer_data['进线时间日期'].unique():
                date_data = customer_data[customer_data['进线时间日期'] == date]['进线-入库时间差']
                interval_count = {'<0': 0}
                for i in range(15):
                    interval_count[f'{i}-{i+1}'] = 0
                interval_count['>14'] = 0

                for value in date_data:
                    if value < 0:
                        interval_count['<0'] += 1
                    elif 0 <= value <= 14:
                        interval_count[f'{int(value)}-{int(value)+1}'] += 1
                    else:
                        interval_count['>14'] += 1

                total_count = sum(interval_count.values())
                date_count[date] = {'interval_count': interval_count, 'total_count': total_count}

            customer_date_count[customer] = date_count

        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        generated_charts = []  # 用于收集生成的图表文件路径
        generated_pareto_charts = []  # 用于收集生成的帕累托图文件路径

        # 为帕累托图设置颜色列表
        colors = ['r', 'g', 'b', 'c', 'm', 'y', 'k', 'orange', 'purple', 'pink']
        color_index = 0

        for customer in customers_with_counts_gt1:
            date_count = len(customer_date_count[customer])
            if date_count > 1:
                # 生成时间差计数图
                fig, ax = plt.subplots(figsize=(14, 10))
                dates = list(customer_date_count[customer].keys())
                intervals = list(customer_date_count[customer][dates[0]]['interval_count'].keys())

                for date in dates:
                    counts = list(customer_date_count[customer][date]['interval_count'].values())
                    ax.plot(intervals, counts, label=f"{date} - {customer_date_count[customer][date]['total_count']}", 
                           marker='o', linestyle='-')
                    
                    # 添加数值标注
                    for i, count in enumerate(counts):
                        ax.text(i, count + 0.2, str(count), ha='center', fontsize=9)

                ax.legend(fontsize=10, loc='upper right')
                ax.set_title(f'{customer} -进线-入库时间差计数', fontsize=16, fontweight='bold')
                ax.set_xlabel('进线-入库时间差', fontsize=14)
                ax.set_ylabel('计数', fontsize=14)
                ax.grid(True, linestyle='--', alpha=0.7)
                ax.set_facecolor('#f5f5f5')
                ax.tick_params(axis='both', labelsize=12)

                output_path = os.path.join(output_dir, f'{customer}_time_diff_count.png')
                plt.savefig(output_path, dpi=300, bbox_inches='tight')
                plt.close()  
                generated_charts.append(output_path)  

                # 重新生成帕累托图，使用双Y轴
                fig_pareto, ax_pareto = plt.subplots(figsize=(16, 10))
                ax_pareto_right = ax_pareto.twinx()  # 创建右Y轴
                positions = np.arange(len(intervals))  # 用于错开标注
                label_offset = 0.1  # 标签偏移量
                
                for idx, date in enumerate(dates):
                    # 获取颜色（循环使用颜色列表）
                    color = colors[color_index % len(colors)]
                    color_index += 1
                    
                    counts = list(customer_date_count[customer][date]['interval_count'].values())
                    percentages = np.cumsum(counts) / sum(counts) * 100
                    
                    # 绘制柱状图
                    bars = ax_pareto.bar(intervals, counts, alpha=0.6, label=f"{date} - 计数", color=color)
                    
                    # 添加柱状图数值标注
                    ax_pareto.bar_label(bars, padding=3, fontsize=9)
                    
                    # 绘制折线图在右Y轴
                    line, = ax_pareto_right.plot(intervals, percentages, marker='o', linestyle='-', linewidth=2, 
                                                label=f"{date} - 累积百分比", color=color)
                    
                    # 添加折线图数值标注，错开位置
                    for i, (pct, pos) in enumerate(zip(percentages, positions)):
                        ax_pareto_right.text(pos, pct + label_offset, f"{pct:.1f}%", ha='center', fontsize=9)
                        label_offset = -label_offset if abs(label_offset) < 0.5 else -0.1  # 自动调整偏移方向

                # 设置左Y轴（计数）
                ax_pareto.set_ylabel('计数', fontsize=14, color='blue')
                ax_pareto.tick_params(axis='y', labelcolor='blue')
                
                # 设置右Y轴（百分比）
                ax_pareto_right.set_ylabel('累积百分比', fontsize=14, color='red')
                ax_pareto_right.tick_params(axis='y', labelcolor='red')
                ax_pareto_right.set_ylim(0, 100)  # 确保百分比范围为0-100%
                
                # 统一标题和坐标轴标签
                ax_pareto.set_title(f'{customer} - 进线-入库时间差帕累托图', fontsize=16, fontweight='bold')
                ax_pareto.set_xlabel('进线-入库时间差', fontsize=14)
                
                # 合并图例并放置在图表外部
                handles1, labels1 = ax_pareto.get_legend_handles_labels()
                handles2, labels2 = ax_pareto_right.get_legend_handles_labels()
                fig_pareto.legend(handles1 + handles2, labels1 + labels2, 
                                  loc='upper center', bbox_to_anchor=(0.5, -0.05),
                                  fancybox=True, shadow=True, ncol=2)
                
                # 其他设置
                ax_pareto.grid(True, linestyle='--', alpha=0.7)
                ax_pareto.set_facecolor('#f5f5f5')
                ax_pareto.tick_params(axis='both', labelsize=12)

                pareto_output_path = os.path.join(output_dir, f'{customer}_pareto.png')
                plt.savefig(pareto_output_path, dpi=300, bbox_inches='tight')
                plt.close()  
                generated_pareto_charts.append(pareto_output_path)

        # 返回所有生成的图表文件路径
        if generated_charts and generated_pareto_charts:
            return {
                'success': True,
                'message': "操作成功完成！已生成以下图表文件：\n" + "\n".join(generated_charts + generated_pareto_charts),
                'generated_charts': generated_charts + generated_pareto_charts
            }
        elif generated_charts:
            return {
                'success': True,
                'message': "操作成功完成！已生成以下计数图表文件：\n" + "\n".join(generated_charts),
                'generated_charts': generated_charts
            }
        else:
            return {'success': True, 'message': "操作成功完成！未生成图表文件。"}

    except Exception as e:
        return {'success': False, 'message': str(e)}

# 如果直接运行此脚本（而非被导入），则可以从命令行获取参数并调用 main 函数
if __name__ == "__main__":
    direct_table_path = input("请输入直接表格数据的路径（如果没有，输入n）：").strip('"')
    output_dir = input("请输入输出文件夹的绝对路径：").strip('"')
    file_paths = None  

    if direct_table_path.lower() == 'n':
        file_paths = input("请输入多个表格文件的路径（用逗号分隔）：").strip('"')
    
    result = main(direct_table_path, output_dir, file_paths)
    print(result['message'])