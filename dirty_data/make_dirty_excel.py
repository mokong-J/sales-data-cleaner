# -*- coding: utf-8 -*-
"""
脏数据生成器 —— 项目 1 专用
================================
生成 12 份「故意很脏」的销售 Excel，用于练习多文件合并清洗。

运行：python make_dirty_excel.py
输出：./dirty_data/ 目录下 12 个文件

刻意制造的坑（每一条都对应真实接单场景）：
  1. 字段名不一致  —— 「销售额」/「销售金额」/「amount」混用
  2. 数字列混文本  —— 「暂无」「-」「N/A」混在金额里
  3. 日期格式混乱  —— 三种格式混用（含中文日期）
  4. 重复行        —— 每月随机插入重复记录
  5. 全角半角混用  —— 地区名有的全角括号有的半角
  6. 前后空格      —— 产品名带首尾空格
  7. 编码差异      —— 其中一份是 GBK 编码 CSV，一份是带 BOM 的 CSV
  8. 空值          —— 随机缺失
  9. 额外空行      —— 文件末尾混入空行
 10. 字段顺序不同  —— 部分文件列顺序被打乱
"""

import os
import random
import pandas as pd

random.seed(42)  # 固定随机种子，保证每次生成的数据一致，方便调试

OUT_DIR = "dirty_data"
os.makedirs(OUT_DIR, exist_ok=True)

REGIONS = ["华东", "华北", "华南", "西南", "西北", "东北"]
PRODUCTS = ["笔记本电脑", "显示器", "键盘", "鼠标", "打印机", "服务器"]

# 字段名不一致的三种写法（对应真实场景里不同人导出的表）
FIELD_VARIANTS = {
    "date": ["日期", "日期", "销售日期"],
    "region": ["地区", "销售区域", "region"],
    "product": ["产品", "商品名称", "product"],
    "amount": ["销售额", "销售金额", "amount"],
    "qty": ["数量", "销量", "quantity"],
}

# 三种日期格式，按月份轮换（模拟不同人导出的习惯）
DATE_FORMATS = [
    "%Y-%m-%d",      # 2024-01-15
    "%Y/%m/%d",      # 2024/01/15
    "%Y年%m月%d日",   # 2024年01月15日
]


def make_month_data(year, month):
    """生成某个月的销售数据（干净的原始数据）"""
    rows = []
    for _ in range(random.randint(80, 120)):
        day = random.randint(1, 28)
        rows.append({
            "date": f"{year}-{month:02d}-{day:02d}",
            "region": random.choice(REGIONS),
            "product": random.choice(PRODUCTS),
            "amount": round(random.uniform(1000, 50000), 2),
            "qty": random.randint(1, 50),
        })
    return pd.DataFrame(rows)


def apply_dirty(df, month_idx):
    """按月份施加不同的脏化策略"""
    # --- 1. 字段名不一致：每 3 个月换一套，循环使用 ---
    variant_group = (month_idx // 3) % 3
    rename_map = {k: v[variant_group] for k, v in FIELD_VARIANTS.items()}
    df = df.rename(columns=rename_map)
    amount_col = rename_map["amount"]

    # 金额列先转成 object，避免写入文本时触发 dtype 警告
    df[amount_col] = df[amount_col].astype(object)

    # --- 2. 数字列混入文本 ---
    if month_idx >= 6:  # 后 6 个月开始出现
        dirty_tokens = ["暂无", "-", "N/A", "待定", ""]
        idx = df.sample(frac=0.08, random_state=month_idx).index
        for i in idx:
            df.at[i, amount_col] = random.choice(dirty_tokens)

    # --- 3. 日期格式混乱：三种格式轮换 ---
    date_col = rename_map["date"]
    fmt = DATE_FORMATS[month_idx % 3]
    df[date_col] = pd.to_datetime(df[date_col]).dt.strftime(fmt)

    # --- 4. 重复行 ---
    n_dup = random.randint(3, 8)
    dups = df.sample(n=n_dup, random_state=month_idx + 100)
    df = pd.concat([df, dups], ignore_index=True)

    # --- 5 & 6. 全角半角混用 + 前后空格 ---
    region_col = rename_map["region"]
    product_col = rename_map["product"]
    if month_idx % 2 == 0:
        # 地区名加全角括号备注
        df[region_col] = df[region_col].apply(
            lambda x: f"{x}（大区）" if random.random() < 0.3 else x
        )
    if month_idx % 3 == 0:
        # 产品名带首尾空格和不可见字符
        df[product_col] = df[product_col].apply(lambda x: f"  {x} ")

    # --- 8. 随机缺失值 ---
    if month_idx >= 4:
        missing_idx = df.sample(frac=0.05, random_state=month_idx + 200).index
        df.loc[missing_idx, rename_map["qty"]] = None

    # --- 10. 字段顺序打乱（部分文件） ---
    if month_idx % 4 == 3:
        df = df[list(reversed(df.columns))]

    return df


def main():
    print("开始生成 12 份脏数据...")
    print("-" * 60)

    generated = []

    for month_idx in range(12):
        month = month_idx + 1
        year = 2024 if month <= 12 else 2025

        df = make_month_data(year, month)
        df = apply_dirty(df, month_idx)

        # 文件名也不统一（真实场景常见）
        if month_idx < 4:
            fname = f"销售数据_{month:02d}月.xlsx"
        elif month_idx < 8:
            fname = f"sales_{month:02d}.xlsx"
        else:
            fname = f"{month}月份销售.xlsx"

        fpath = os.path.join(OUT_DIR, fname)
        df.to_excel(fpath, index=False)
        generated.append(fpath)
        print(f"  ✓ {fname:<24} {len(df):>4} 行 | 字段: {list(df.columns)}")

    # --- 7. 编码差异：额外造两份 CSV 制造编码坑 ---
    extra = make_month_data(2024, 12)
    extra = apply_dirty(extra, 11)

    # GBK 编码
    gbk_path = os.path.join(OUT_DIR, "销售数据_补充_GBK.csv")
    extra.to_csv(gbk_path, index=False, encoding="gbk")
    generated.append(gbk_path)
    print(f"  ✓ {'销售数据_补充_GBK.csv':<24} {len(extra):>4} 行 | 编码: GBK ← 坑点")

    # 带 BOM 的 UTF-8
    bom_path = os.path.join(OUT_DIR, "销售数据_补充_BOM.csv")
    extra.to_csv(bom_path, index=False, encoding="utf-8-sig")
    generated.append(bom_path)
    print(f"  ✓ {'销售数据_补充_BOM.csv':<24} {len(extra):>4} 行 | 编码: UTF-8-BOM ← 坑点")

    print("-" * 60)
    print(f"生成完毕，共 {len(generated)} 个文件，位于 ./{OUT_DIR}/")
    print()
    print("你的任务：写一个脚本把这 12+2 份文件合并成一张干净的总表")
    print("提示：先 print 每份文件的 columns 和 shape，看看坑在哪")


if __name__ == "__main__":
    main()
