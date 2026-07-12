from pathlib import Path

import pandas as pd


project_dir = Path(__file__).resolve().parent.parent
data_dir = project_dir / "02_清洗数据"

excel_files = list(data_dir.glob("*.xlsx"))

if not excel_files:
    raise FileNotFoundError("在 02_清洗数据 中没有找到 xlsx 文件")

file_path = excel_files[0]

print("=" * 60)
print("正在读取文件：", file_path.name)
print("数据较大，读取可能需要一两分钟，请耐心等待。")
print("=" * 60)

df = pd.read_excel(file_path, sheet_name=0)

print("\n【数据规模】")
print("总行数：", len(df))
print("总列数：", len(df.columns))

print("\n【字段名称】")
print(df.columns.tolist())

print("\n【各字段数据类型】")
print(df.dtypes)

print("\n【缺失值数量】")
print(df.isna().sum())

print("\n【完全重复的行数】")
print(df.duplicated().sum())

invoice_text = df["Invoice"].astype(str)
cancelled_count = invoice_text.str.startswith("C", na=False).sum()

print("\n【订单异常检查】")
print("取消订单明细行数：", cancelled_count)
print("数量小于或等于 0 的行数：", (df["Quantity"] <= 0).sum())
print("价格小于或等于 0 的行数：", (df["Price"] <= 0).sum())

print("\n【时间范围】")
print("最早交易时间：", df["InvoiceDate"].min())
print("最晚交易时间：", df["InvoiceDate"].max())

print("\n【基础业务规模】")
print("不同订单编号数：", df["Invoice"].nunique())
print("不同商品数：", df["StockCode"].nunique())
print("不同客户数：", df["Customer ID"].nunique())
print("不同国家或地区数：", df["Country"].nunique())

print("\n数据检查完成。")
