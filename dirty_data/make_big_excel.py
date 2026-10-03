# -*- coding: utf-8 -*-
"""
大文件生成器 —— 覆盖生成 30 万行的 10月份销售.xlsx，用于测试合并脚本性能
运行：python make_big_excel.py
还原：python make_dirty_excel.py（恢复全部 12+2 份原始小文件）
"""

import os
import random

import pandas as pd

random.seed(42)  # 与 make_dirty_excel.py 保持一致，保证数据可复现

OUT_DIR = os.path.dirname(os.path.abspath(__file__))
N_ROWS = 300_000

REGIONS = ["华东", "华北", "华南", "西南", "西北", "东北"]
PRODUCTS = ["笔记本电脑", "显示器", "键盘", "鼠标", "打印机", "服务器"]
DIRTY_TOKENS = ["暂无", "-", "N/A", "待定", ""]


def main():
    print(f"开始生成 {N_ROWS} 行数据...")
    df = pd.DataFrame({
        # 10 月份原文件为 %Y-%m-%d 格式
        "日期": [f"2024-10-{random.randint(1, 28):02d}" for _ in range(N_ROWS)],
        "地区": [random.choice(REGIONS) for _ in range(N_ROWS)],
        # 产品名带首尾空格，复刻原 10 月份文件的坑
        "产品": [f"  {random.choice(PRODUCTS)} " for _ in range(N_ROWS)],
        # 金额先转 object，避免写入文本时触发 dtype 警告
        "销售额": [round(random.uniform(1000, 50000), 2) for _ in range(N_ROWS)],
        "数量": [random.randint(1, 50) for _ in range(N_ROWS)],
    })
    df["销售额"] = df["销售额"].astype(object)

    # 金额列混入文本脏数据，比例与 make_dirty_excel.py 的策略一致
    dirty_idx = df.sample(frac=0.08, random_state=9).index
    for i in dirty_idx:
        df.at[i, "销售额"] = random.choice(DIRTY_TOKENS)

    # 数量列随机缺失
    missing_idx = df.sample(frac=0.05, random_state=209).index
    df.loc[missing_idx, "数量"] = None

    fpath = os.path.join(OUT_DIR, "10月份销售.xlsx")
    df.to_excel(fpath, index=False)
    size_mb = os.path.getsize(fpath) / 1024 / 1024
    print(f"已生成 {fpath}")
    print(f"共 {len(df)} 行，文件大小 {size_mb:.1f} MB")


if __name__ == "__main__":
    main()
