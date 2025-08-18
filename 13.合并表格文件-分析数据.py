#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
merge_analyze_v2.py
交互式合并 CSV 并按需做数值/非数值分析（升级版）
python merge_analyze_v2.py
"""

import os
import sys
from pathlib import Path

import pandas as pd


# ---------- 工具函数 ----------
def read_csv(prompt):
    """交互式读取 CSV，确保文件存在并可读。"""
    while True:
        path = input(prompt).strip().strip('"')
        if not os.path.isfile(path):
            print("❌ 文件不存在，请重新输入：")
            continue
        try:
            df = pd.read_csv(path, dtype=str)  # 先全部按字符串读
            print(f"✅ 成功读取 {path}，共 {len(df)} 行。")
            return df, Path(path)
        except Exception as e:
            print(f"❌ 读取失败：{e}\n请重新输入：")


def ask_column(df, name):
    """询问并校验列名。"""
    print(f"{name} 的列有：{list(df.columns)}")
    while True:
        col = input(f"请输入用于匹配的列名（{name}）：").strip()
        if col in df.columns:
            return col
        print("❌ 列名不存在，请重新输入：")


def merge_tables(df_a, key_a, df_b, key_b):
    """以指定列为键做外连接，保留两表其余列。"""
    merged = pd.merge(
        df_a,
        df_b,
        left_on=key_a,
        right_on=key_b,
        how="outer",
        suffixes=("_A", "_B"),
    )
    print("✅ 合并完成。")
    return merged


def ask_analysis_type():
    choices = {
        "1": "数值分析",
        "2": "非数值分析",
        "3": "不进行分析",
        "4": "两者都分析",
    }
    prompt = "\n请选择后续分析：\n" + "\n".join([f"{k}. {v}" for k, v in choices.items()]) + "\n请输入对应数字："
    while True:
        ans = input(prompt).strip()
        if ans in choices:
            return ans
        print("❌ 输入无效，请重新选择：")


# ---------- 分析函数 ----------
def numeric_bucket(val):
    """数值 → 区间名称"""
    if pd.isna(val):
        return "未入库"
    if val < 0:
        return "未入库"
    elif 0 <= val < 3:
        return "入库0-3天"
    elif 3 <= val < 5:
        return "入库3-5天"
    else:
        return "入库超5天"


def numeric_analysis(df, col):
    """数值列分段统计，返回 DataFrame（含“数值区间”列）"""
    try:
        s = pd.to_numeric(df[col], errors="coerce")
    except Exception as e:
        print(f"⚠️ 无法将列 {col} 转为数值：{e}")
        return pd.DataFrame()

    buckets = s.apply(numeric_bucket)
    summary = (
        buckets.value_counts()
        .rename("数量")
        .to_frame()
        .assign(
            数值区间=lambda x: x.index,
            占比=lambda x: (x["数量"] / len(s)).apply(lambda p: f"{p:.2%}"),
        )[["数值区间", "数量", "占比"]]
    )
    summary.loc["总计"] = ["总计", len(s), "100.00%"]
    return summary


def non_numeric_analysis(df, col):
    """非数值列分组统计，返回 DataFrame"""
    s = df[col].astype(str).replace({"nan": None, "None": None, "": None})
    summary = (
        s.value_counts(dropna=False)
        .rename("数量")
        .to_frame()
        .assign(占比=lambda x: (x["数量"] / len(s)).apply(lambda p: f"{p:.2%}"))
    )
    summary.index.name = col
    summary = summary.reset_index()
    summary.loc[len(summary)] = ["总计", len(s), "100.00%"]
    return summary


# ---------- 主流程 ----------
def main():
    print("=== 表格合并 & 分析工具（升级版） ===\n")

    # 1. 读取 A、B 表
    df_a, path_a = read_csv("请输入表 A 的 CSV 文件路径：")
    df_b, path_b = read_csv("请输入表 B 的 CSV 文件路径：")

    # 2. 合并
    key_a = ask_column(df_a, "表 A")
    key_b = ask_column(df_b, "表 B")
    merged = merge_tables(df_a, key_a, df_b, key_b)

    # 3. 保存合并结果（CSV）
    out_dir = path_a.parent
    merge_csv = out_dir / "合并结果.csv"
    merged.to_csv(merge_csv, index=False, encoding="utf-8-sig")
    print(f"✅ 合并结果已保存为：{merge_csv.resolve()}")

    # 4. 询问分析类型
    analysis_type = ask_analysis_type()
    if analysis_type == "3":
        print("已选择不进行分析，程序结束。")
        return

    # 5. 收集分析结果
    results = {}  # sheet名 -> DataFrame
    numeric_col, non_numeric_col = None, None

    # 数值分析
    if analysis_type in {"1", "4"}:
        while True:
            numeric_col = input("请输入需做数值分析的列名：").strip()
            if numeric_col in merged.columns:
                break
            print("❌ 列名不存在，请重新输入：")
        results["数值分析"] = numeric_analysis(merged, numeric_col)

    # 非数值分析（独立）
    if analysis_type in {"2", "4"}:
        while True:
            non_numeric_col = input("请输入需做非数值分析的列名：").strip()
            if non_numeric_col in merged.columns:
                break
            print("❌ 列名不存在，请重新输入：")
        results["非数值分析"] = non_numeric_analysis(merged, non_numeric_col)

    # 6. 如果同时做两种分析 → 额外做“按数值区间分组的非数值分析”
    if analysis_type == "4" and numeric_col and non_numeric_col:
        try:
            merged["__数值区间"] = pd.to_numeric(
                merged[numeric_col], errors="coerce"
            ).apply(numeric_bucket)
        except Exception:
            merged["__数值区间"] = "未入库"

        for zone, sub in merged.groupby("__数值区间"):
            sheet_name = f"非数值_{zone}"
            results[sheet_name] = non_numeric_analysis(sub, non_numeric_col)

    # 7. 输出 Excel
    out_excel = out_dir / "数据分析.xlsx"
    with pd.ExcelWriter(out_excel, engine="openpyxl") as writer:
        for sheet, df_res in results.items():
            df_res.to_excel(writer, sheet_name=sheet, index=False)
    print(f"✅ 数据分析结果已保存为：{out_excel.resolve()}\n")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n用户中断，程序退出。")
        sys.exit(0)