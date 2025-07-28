import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os
import io
from collections import Counter

PARAM_PROMPTS = {
    'input_path': "请输入“客户明细-时间差值明细文件”路径(注：csv格式)：",
    'output_dir': "请输入输出文件夹的绝对路径：",
    'date_prefix': '请输入输出文件的日期前缀：'
}

plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False

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
        if not input_path.lower().endswith('.csv'):
            raise ValueError("不支持的文件格式，请提供CSV文件")
    
        df = pd.read_csv(input_path)

        selected_data = df[['客户名称', '进线-入库时间差']]
        customers = selected_data['客户名称'].unique()

        customer_count = {}
        for customer in customers:
            customer_data = selected_data[selected_data['客户名称'] == customer]['进线-入库时间差']
            interval_count = {'<0': 0}
            for i in range(14):
                interval_count[f'{i}-{i+1}'] = 0
            interval_count['>14'] = 0

            for value in customer_data:
                if value < 0:
                    interval_count['<0'] += 1
                elif 0 <= value <= 14:
                    interval_count[f'{int(value)}-{int(value)+1}'] += 1
                else:
                    interval_count['>14'] += 1

            customer_count[customer] = interval_count

        # 为所有客户生成总览图
        fig_total, ax_total = plt.subplots(figsize=(14, 10))
        colors_total = plt.cm.get_cmap('tab20c', len(customers))

        for i, customer in enumerate(customers):
            intervals = list(customer_count[customer].keys())
            counts = list(customer_count[customer].values())
            total_count = sum(counts)
            ax_total.plot(intervals, counts, label=f"{customer} - {total_count}", 
                          color=colors_total(i), linewidth=2, linestyle='-', marker='o')

        # 设置图例位置为外部右侧
        ax_total.legend(bbox_to_anchor=(1.05, 1), loc='upper left', fontsize=10)
        ax_total.set_title(f'{date_prefix}-所有客户进线-入库时间差计数总览', fontsize=16, fontweight='bold')
        ax_total.set_xlabel('进线-入库时间差', fontsize=14)
        ax_total.set_ylabel('计数', fontsize=14)
        ax_total.grid(True, linestyle='--', alpha=0.7)
        ax_total.set_facecolor('#f5f5f5')
        ax_total.tick_params(axis='both', labelsize=12)

        plt.tight_layout()
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        total_output_path = os.path.join(output_dir, '所有客户进线-入库时间差计数总览.png')
        plt.savefig(total_output_path, dpi=300, bbox_inches='tight')
        
        # 将总览图转换为字节流
        total_image_buffer = io.BytesIO()
        plt.savefig(total_image_buffer, format='png')
        plt.close()  # 关闭图表以释放资源

        # 为每个客户生成单独的图表
        for customer in customers:
            fig_individual, ax_individual = plt.subplots(figsize=(12, 8))
            intervals = list(customer_count[customer].keys())
            counts = list(customer_count[customer].values())
            
            total_count = sum(counts)
            ax_individual.plot(intervals, counts, marker='o', linestyle='-', linewidth=2, color='purple')
            
            # 在图上显示数值
            for i, count in enumerate(counts):
                ax_individual.text(i, count + 0.5, str(count), ha='center', fontsize=16)
            
            ax_individual.set_title(f' {customer} -{date_prefix}-进线-入库时间差计数：{total_count}', fontsize=16, fontweight='bold')
            ax_individual.set_xlabel('进线-入库时间差', fontsize=14)
            ax_individual.set_ylabel('计数', fontsize=18)
            ax_individual.grid(True, linestyle='--', alpha=0.7)
            ax_individual.set_facecolor('#f5f5f5')
            ax_individual.tick_params(axis='both', labelsize=12)
            
            plt.tight_layout()
            customer_output_path = os.path.join(output_dir, f'客户_{customer}_进线-入库时间差计数.png')
            plt.savefig(customer_output_path, dpi=300, bbox_inches='tight')
            plt.close()  # 关闭图表以释放资源

        # 生成帕累托图
        def generate_pareto_chart(customer, intervals, counts, output_path):
            # 计算累积百分比
            percentages = np.cumsum(counts) / sum(counts) * 100
            fig, ax1 = plt.subplots(figsize=(12, 8))

            # 绘制柱状图
            bars = ax1.bar(intervals, counts, color='skyblue', label='计数')
            ax1.set_xlabel('进线-入库时间差', fontsize=14)
            ax1.set_ylabel('计数', fontsize=14, color='b')
            ax1.tick_params(axis='y', labelcolor='b')
            ax1.set_facecolor('#f5f5f5')
            ax1.grid(True, linestyle='--', alpha=0.7)
            
            # 在柱状图上显示数值
            for bar in bars:
                height = bar.get_height()
                ax1.text(bar.get_x() + bar.get_width()/2., height + 0.5,
                        f'{height}', ha='center', va='bottom', fontsize=12)

            # 绘制累积百分比折线图（调整颜色为紫色）
            ax2 = ax1.twinx()
            line, = ax2.plot(intervals, percentages, color='red', marker='o', linestyle='-', linewidth=2, label='累积百分比')
            
            # 在折线图上显示数值（调整字体大小）
            for i, percentage in enumerate(percentages):
                ax2.text(i, percentage + 1, f'{percentage:.1f}%', ha='center', va='bottom', fontsize=12)
            
            # 添加图例（放在图外）
            fig.legend([bars, line], ['计数', '累积百分比'], 
                      loc='upper left', bbox_to_anchor=(0.1, 0.9), fontsize=10)

            # 设置标题
            ax1.set_title(f'{customer} - {date_prefix} - 帕累托图', fontsize=16, fontweight='bold')

            plt.tight_layout()
            # 保存图表
            plt.savefig(output_path, dpi=300, bbox_inches='tight')
            plt.close()

        # 生成所有客户的帕累托图总览
        fig_pareto_total, ax_pareto_total = plt.subplots(figsize=(14, 10))
        for i, customer in enumerate(customers):
            intervals = list(customer_count[customer].keys())
            counts = list(customer_count[customer].values())
            percentages = np.cumsum(counts) / sum(counts) * 100

            ax_pareto_total.plot(intervals, percentages, label=f"{customer}",
                                 color=plt.cm.get_cmap('tab20c', len(customers))(i),
                                 linewidth=2, linestyle='-', marker='o')

        ax_pareto_total.set_title(f'{date_prefix}-所有客户进线-入库时间差帕累托图总览', fontsize=16, fontweight='bold')
        ax_pareto_total.set_xlabel('进线-入库时间差', fontsize=14)
        ax_pareto_total.set_ylabel('累积百分比', fontsize=14)
        ax_pareto_total.legend(bbox_to_anchor=(1.05, 1), loc='upper left', fontsize=10)
        ax_pareto_total.grid(True, linestyle='--', alpha=0.7)
        ax_pareto_total.set_facecolor('#f5f5f5')
        ax_pareto_total.tick_params(axis='both', labelsize=12)

        plt.tight_layout()
        pareto_total_output_path = os.path.join(output_dir, '所有客户进线-入库时间差帕累托图总览.png')
        plt.savefig(pareto_total_output_path, dpi=300, bbox_inches='tight')
        plt.close()

        # 为每个客户生成单独的帕累托图
        for customer in customers:
            intervals = list(customer_count[customer].keys())
            counts = list(customer_count[customer].values())
            generate_pareto_chart(customer, intervals, counts, 
                                 os.path.join(output_dir, f'客户_{customer}_进线-入库时间差帕累托图.png'))

        # 返回总览图数据（客户单独图保存到文件夹中）
        return {
            'success': True,
            'message': f"操作成功完成！已生成总览图和{len(customers)}个客户单独图表，以及对应的帕累托图",
            'total_image_data': total_image_buffer.getvalue()
        }
    
    except Exception as e:
        # 返回失败状态和错误信息
        return {'success': False, 'message': str(e)}

# 如果直接运行此脚本（而非被导入），则可以从命令行获取参数并调用 main 函数
if __name__ == "__main__":
    input_path = input("请输入“客户明细-时间差值明细文件”路径(注：csv格式)：").strip('"')
    output_dir = input("请输入输出文件夹的绝对路径：").strip('"')
    date_prefix = input('请输入输出文件的日期前缀：').strip('"')
    result = main(input_path, output_dir, date_prefix)
    print(result['message'])