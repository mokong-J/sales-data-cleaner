# -*- coding: utf-8 -*-
"""
项目 1 主脚本：多文件销售数据合并清洗
流程：读取 dirty_data 下全部 Excel/CSV -> 字段名统一 -> 合并
      -> 日期/数值清洗 -> 缺失剔除 -> 去重 -> 输出明细 + 透视表 + 数据质量报告
"""

import re
import sys
import time
from pathlib import Path

import pandas as pd
from charset_normalizer import from_path


def norm(s):
    """
    标准化字段名称：去首尾空格、移除空格/下划线/短横线、删除括号及括号内内容、转小写
    例：「金额(元)」->「金额」、「华东（大区）」->「华东」；转小写让列名匹配不区分大小写
    :param s: 输入字符串
    :return: 标准化后的字符串
    """
    s = s.strip()
    s = re.sub(r"[\s_-]", "", s)
    s = re.sub(r"[（(].*?[)）]", "", s)
    s = s.lower()
    return s


start_time = time.time()

# 标准列名，用于合并时的重命名与输出列顺序
std_cols = ["Date", "Region", "Product", "Amount", "Quantity"]

# 字段别名表，处理不同文件字段名不一致的情况（标准名 -> 别名列表）
alias = {
    "Date": [
        "日期",
        "销售日期",
        "日期时间",
        "订单日期",
        "统计日期",
        "date",
        "order_date",
        "sale_date",
    ],
    "Region": [
        "地区",
        "销售区域",
        "区域",
        "大区",
        "省份",
        "城市",
        "region",
        "area",
        "city",
        "province",
    ],
    "Product": [
        "产品",
        "商品名称",
        "产品名称",
        "商品",
        "品类",
        "product",
        "item",
        "goods",
        "sku",
    ],
    "Amount": [
        "销售额",
        "销售金额",
        "金额",
        "金额(元)",
        "营业额",
        "amount",
        "sales",
        "sale_amount",
        "revenue",
        "gmv",
    ],
    "Quantity": [
        "数量",
        "销量",
        "销售数量",
        "件数",
        "qty",
        "quantity",
        "count",
        "units",
        "volume",
    ],
}

# 字段别名统一化，用于匹配标准化后的文件字段名
alias = {key: [norm(value) for value in values] for key, values in alias.items()}

file_dir = Path(__file__).parent / "dirty_data"
if not file_dir.exists():
    sys.exit(f"错误：{file_dir} 目录不存在")

failed_files = []  # 读取失败或空文件：只提示跳过，不中断整批
total_rows = 0  # 成功读取文件的总行数，供最后的数字自洽校验
files_count = 0
df_list = []  # 先收集再一次性拼接，避免循环内反复 concat 大表（越拼越慢）
for file_path in [*file_dir.glob("*.xls*"), *file_dir.glob("*.csv")]:
    # Excel 打开文件时的临时副本以 ~$ 开头，不是数据，直接跳过
    if file_path.name.startswith("~$"):
        continue
    try:
        if file_path.suffix == ".csv":
            # CSV 编码未知，先自动检测（可覆盖 GBK / UTF-8 / UTF-8-BOM）
            encoding = from_path(file_path).best().encoding
            df = pd.read_csv(file_path, encoding=encoding)
        else:
            df = pd.read_excel(file_path, engine="calamine")
    except Exception as e:
        print(f"错误：{file_path} 读取失败，错误信息：{e}")
        failed_files.append(file_path.name)
        continue
    if df.empty:
        print(f"错误：{file_path.name} 为空，跳过")
        failed_files.append(file_path.name)
        continue
    total_rows += df.shape[0]

    # 统一列名：先标准化列名，再按别名表重命名。
    # 每个标准名只取第一个命中的列（按字段名匹配、不按列位置匹配，
    # 列顺序打乱也不受影响）；没命中别名的列会在 reindex 时丢弃
    df.columns = [norm(col) for col in df.columns]
    rename_map = {}
    for key, values in alias.items():
        for col in df.columns:
            if col in values:
                rename_map[col] = key
                break
    df = df.rename(columns=rename_map)

    # 字段缺失明说但不中断：缺的标准列由 reindex 补 NaN，后续统一按缺失行剔除
    missing_cols = [col for col in std_cols if col not in df.columns]
    if missing_cols:
        print(f"警告：{file_path.name} 缺少字段：{','.join(missing_cols)}")

    # reindex：只保留 5 个标准字段（多余列丢弃、缺失列补 NaN）
    df = df.reindex(columns=std_cols)

    df_list.append(df)
    files_count += 1

if not df_list:
    sys.exit("没有成功获取文件，终止处理")

# 一次性拼接并重置索引，避免行号重复
merge_df = pd.concat(df_list, ignore_index=True)
total_rows_merge = merge_df.shape[0]
print(
    f"成功合并 {files_count} 个文件，共 {total_rows_merge} 行数据, 读取失败 {len(failed_files)} 个文件"
)

if failed_files:
    print(f"读取失败文件清单：{','.join(failed_files)}")

# 日期格式统一化：先把「年 / 月 / . / /」替换为「-」并去掉「日」，
# 再按固定格式解析，解析失败的置为 NaT（不中断，数量计入质量报告）
merge_df["Date"] = pd.to_datetime(
    merge_df["Date"]
    .astype(str)
    .str.strip()
    .str.replace(r"[年月./]", "-", regex=True)
    .str.replace(r"日", "", regex=True),
    errors="coerce",
    format="%Y-%m-%d",
)
date_bad_rows = int(merge_df["Date"].isna().sum())
print(f"日期格式统一化：共 {date_bad_rows} 个日期格式错误")

# 数值列清洗：混入的「暂无 / - / N/A」等文本统一转为 NaN
for col in ["Amount", "Quantity"]:
    merge_df[col] = pd.to_numeric(merge_df[col], errors="coerce")
amount_bad_rows = int(merge_df["Amount"].isna().sum())
quantity_bad_rows = int(merge_df["Quantity"].isna().sum())
print(f"Amount 有 {amount_bad_rows} 行无法解析为数字")
print(f"Quantity 有 {quantity_bad_rows} 行无法解析为数字")

# 只剔除关键字段缺失的行（任一标准列缺失即剔除），不做全表 dropna
rows_before_dropna = merge_df.shape[0]
merge_df = merge_df.dropna(subset=std_cols)
rows_after_dropna = merge_df.shape[0]
dropna_rows = rows_before_dropna - rows_after_dropna
print(f"删除缺失值：共 {dropna_rows} 行缺失值(任意列存在缺失值的行)")

# 数量列用可空整数类型，避免输出成 17.0；缺失行已剔除，此处转换不会失败
merge_df["Quantity"] = merge_df["Quantity"].astype("Int64")

# 地区/产品标准化：去空格与括号注记，让「华东（大区）」并入「华东」
# 放在去重之前，归一化后相同的行也能被一并判重
for col in ["Region", "Product"]:
    merge_df[col] = (
        merge_df[col]
        .astype(str)
        .str.strip()
        .str.replace(r"[\s_-]", "", regex=True)
        .str.replace(r"[（(].*?[)）]", "", regex=True)
    )

# 去重：所有标准字段都相同才算重复，保留第一条
merge_df = merge_df.drop_duplicates(subset=std_cols, keep="first")
rows_after_drop_dup = merge_df.shape[0]
rows_drop_dup = rows_after_dropna - rows_after_drop_dup
print(f"去重：共 {rows_drop_dup} 行重复数据")

# 数字自洽硬校验：读入 = 输出 + 缺失剔除 + 重复剔除，对不上直接终止
if total_rows != rows_after_drop_dup + dropna_rows + rows_drop_dup:
    sys.exit(
        f"合并后数据行数与原始数据行数不一致，终止处理，原始数据行数：{total_rows}，"
        f"合并后数据行数：{rows_after_drop_dup}，缺失值行数：{dropna_rows}，去重行数：{rows_drop_dup}"
    )

# ===== 数据质量报告：读入 -> 丢弃原因分布 -> 输出，数字必须对得上 =====
# 说明：日期/金额/数量的解析失败计数互相有重叠（一行可能同时踩多个坑），
#       「缺失行剔除」才是它们的混合合计，报告中已注明避免误读
report = "\n".join(
    [
        "========== 数据质量报告 ==========",
        f"读入行数：{total_rows}",
        "丢弃明细：",
        f"  日期无法解析：{date_bad_rows} 行",
        f"  金额无法解析：{amount_bad_rows} 行（与其他原因有重叠）",
        f"  数量无法解析：{quantity_bad_rows} 行（与其他原因有重叠）",
        f"  缺失行剔除：{dropna_rows} 行（任一关键字段缺失的合计）",
        f"  重复行剔除：{rows_drop_dup} 行",
        f"输出行数：{rows_after_drop_dup}",
        f"自洽校验：{rows_after_drop_dup} + {dropna_rows} + {rows_drop_dup} = {total_rows} ✓",
        "==================================",
    ]
)
print(report)

# 报告落盘到 docs 目录作为交付物；utf-8-sig 保证 Windows 记事本打开不乱码
report_path = Path(__file__).parent / "docs" / "数据质量报告.txt"
report_path.parent.mkdir(parents=True, exist_ok=True)
report_path.write_text(report, encoding="utf-8-sig")

# 透视表：月份 x 地区 的销售额汇总（日期先转 Period 再透视）
pivot = merge_df.pivot_table(
    index=merge_df["Date"].dt.to_period("M"),
    columns="Region",
    values="Amount",
    aggfunc="sum",
)
pivot = pivot.reset_index()

# 明细日期转成 YYYY-MM-DD 字符串，客户在 Excel 里可直接排序和筛选
merge_df["Date"] = merge_df["Date"].dt.strftime("%Y-%m-%d")
try:
    with pd.ExcelWriter(
        "cleaned_sales_data.xlsx",
        engine="xlsxwriter",
    ) as writer:
        merge_df.to_excel(writer, sheet_name="明细", index=False)
        pivot.to_excel(writer, sheet_name="销售额透视表", index=False)
except PermissionError:
    sys.exit("写入失败： cleaned_sales_data.xlsx 被占用，请关闭 Excel 后重新运行")
print(
    f"已输出 cleaned_sales_data.xlsx（明细 {merge_df.shape[0]} 行 + 销售额透视表），"
    f"数据质量报告已写入 docs/{report_path.name}"
)
end_time = time.time()
print(f"处理耗时：{end_time - start_time:.2f} 秒")
