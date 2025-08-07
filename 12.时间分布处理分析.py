#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
hourly_complaint_final.py
"""
import os, glob, chardet
import pandas as pd
import matplotlib.pyplot as plt

plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False

# ---------- 1. 输入 ----------
folder = input('请输入 CSV 所在文件夹路径：').strip().strip('"')
csv_files = glob.glob(os.path.join(folder, '*.csv'))
if not csv_files:
    raise FileNotFoundError('未找到 CSV！')

# ---------- 2. 工具 ----------
def read_with_encoding(path):
    with open(path, 'rb') as f:
        enc = chardet.detect(f.read(100_000))['encoding']
    return pd.read_csv(path, encoding=enc)

# ---------- 3. 询问时间列 ----------
first_df = read_with_encoding(csv_files[0])
print('列名列表：', list(first_df.columns))
time_col = input('请输入“时间”列的名称：').strip()

# ---------- 4. 处理函数 ----------
def hourly_counts(df):
    """返回 Series：index=1-24，name=投诉量"""
    s = pd.Series(0, index=range(1, 25), name='投诉量')
    s.index.name = '小时区间'
    if time_col not in df.columns:
        return None
    df = df.copy()
    df[time_col] = pd.to_datetime(df[time_col], errors='coerce')
    # 向上取整小时
    minute_total = (df[time_col].dt.hour * 60 +
                    df[time_col].dt.minute +
                    df[time_col].dt.second / 60)
    hour_up = ((minute_total + 59) // 60).astype(int).clip(upper=24)
    counts = hour_up.value_counts().reindex(range(1, 25), fill_value=0)
    return counts

def plot_bar(series, title, save_path):
    fig, ax = plt.subplots(figsize=(10, 4))
    bars = ax.bar(series.index.astype(str), series.values, color='#4C72B0')
    ax.set_title(title, fontsize=14)
    ax.set_xlabel('小时区间')
    ax.set_ylabel(series.name)
    ax.set_xticks(series.index.astype(str))
    ax.set_xticklabels([f'{h:02d}:00' for h in series.index], rotation=45)
    for b in bars:
        ax.text(b.get_x() + b.get_width()/2, b.get_height() + 0.2,
                int(b.get_height()), ha='center', va='bottom', fontsize=8)
    plt.tight_layout()
    plt.savefig(save_path, dpi=300)
    plt.close()

# ---------- 5. 主循环 ----------
out_dir = os.path.join(folder, 'hourly_output')
os.makedirs(out_dir, exist_ok=True)

total_counts = pd.Series(0, index=range(1, 25), name='总投诉量')
total_counts.index.name = '小时区间'

for csv_path in csv_files:
    fname = os.path.basename(csv_path)
    df = read_with_encoding(csv_path)
    counts = hourly_counts(df)
    if counts is None:
        print(f'⚠️ {fname} 无时间列，已跳过')
        continue

    # 5-1 每个文件单独保存 csv & png
    single_csv = os.path.join(out_dir, f'输出_{fname}')
    counts.to_csv(single_csv, encoding='utf-8-sig')

    single_png = os.path.join(out_dir, f'{fname}.png')
    plot_bar(counts, f'{fname} 每小时投诉量', single_png)
    
    # **新增提示**
    print(f'✅ {fname} 已处理完成，结果已输出 → {single_csv} & {single_png}')

    # 5-2 累加到总表
    total_counts += counts

# ---------- 6. 汇总 ----------
if total_counts.sum() == 0:
    print('⚠️ 没有任何有效数据，汇总终止')
else:
    # csv
    total_csv = os.path.join(out_dir, '汇总_hourly_total.csv')
    total_counts.to_csv(total_csv, encoding='utf-8-sig')
    # 图
    total_png = os.path.join(out_dir, '汇总_hourly_total.png')
    plot_bar(total_counts, '全部文件 每小时投诉量汇总', total_png)

    print('全部完成！结果位于：', out_dir)