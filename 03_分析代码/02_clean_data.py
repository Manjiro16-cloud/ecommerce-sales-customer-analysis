from pathlib import Path

import pandas as pd


project_dir = Path(__file__).resolve().parent.parent
input_dir = project_dir / "02_清洗数据"
output_dir = project_dir / "02_清洗数据"

input_files = [
    file
    for file in input_dir.glob("*.xlsx")
    if not file.name.startswith("~$")
]

if len(input_files) != 1:
    raise RuntimeError(
        f"02_清洗数据 中应当只有一个 xlsx 文件，实际找到 {len(input_files)} 个"
    )

input_file = input_files[0]

print("=" * 60)
print("开始读取：", input_file.name)
print("=" * 60)

df = pd.read_excel(input_file, sheet_name=0)

original_rows = len(df)
duplicate_rows = int(df.duplicated().sum())

# 统一字段名称，方便后续使用 Python、SQL 和 Excel。
df = df.rename(
    columns={
        "Invoice": "invoice_no",
        "StockCode": "stock_code",
        "Description": "description",
        "Quantity": "quantity",
        "InvoiceDate": "invoice_date",
        "Price": "unit_price",
        "Customer ID": "customer_id",
        "Country": "country",
    }
)

# 删除完全重复的明细行，保留第一次出现的记录。
df = df.drop_duplicates().copy()

# 清理文本字段两端的空格。
df["invoice_no"] = df["invoice_no"].astype(str).str.strip()
df["stock_code"] = df["stock_code"].astype(str).str.strip()
df["description"] = df["description"].astype("string").str.strip()
df["country"] = df["country"].astype("string").str.strip()

# 统一客户编号格式，缺失客户仍保留为空。
df["customer_id"] = pd.to_numeric(
    df["customer_id"], errors="coerce"
).astype("Int64")

# 保证日期、数量和价格是正确的数据类型。
df["invoice_date"] = pd.to_datetime(df["invoice_date"], errors="coerce")
df["quantity"] = pd.to_numeric(df["quantity"], errors="coerce")
df["unit_price"] = pd.to_numeric(df["unit_price"], errors="coerce")

# 标记不同类型的异常或非正常销售记录。
cancelled_mask = df["invoice_no"].str.startswith("C", na=False)
invalid_quantity_mask = df["quantity"].le(0) | df["quantity"].isna()
invalid_price_mask = df["unit_price"].le(0) | df["unit_price"].isna()
missing_description_mask = (
    df["description"].isna()
    | df["description"].eq("")
)
invalid_date_mask = df["invoice_date"].isna()

valid_sales_mask = ~(
    cancelled_mask
    | invalid_quantity_mask
    | invalid_price_mask
    | missing_description_mask
    | invalid_date_mask
)

clean_sales = df.loc[valid_sales_mask].copy()
excluded_records = df.loc[~valid_sales_mask].copy()

# 为被排除的数据记录原因，便于审计和面试说明。
excluded_records["exclusion_reason"] = ""

excluded_records.loc[
    cancelled_mask.loc[excluded_records.index],
    "exclusion_reason",
] += "cancelled_invoice;"

excluded_records.loc[
    invalid_quantity_mask.loc[excluded_records.index],
    "exclusion_reason",
] += "non_positive_quantity;"

excluded_records.loc[
    invalid_price_mask.loc[excluded_records.index],
    "exclusion_reason",
] += "non_positive_price;"

excluded_records.loc[
    missing_description_mask.loc[excluded_records.index],
    "exclusion_reason",
] += "missing_description;"

excluded_records.loc[
    invalid_date_mask.loc[excluded_records.index],
    "exclusion_reason",
] += "invalid_date;"

# 创建后续分析需要的业务字段。
clean_sales["sales_amount"] = (
    clean_sales["quantity"] * clean_sales["unit_price"]
).round(2)

clean_sales["order_date"] = clean_sales["invoice_date"].dt.date
clean_sales["year_month"] = (
    clean_sales["invoice_date"].dt.to_period("M").astype(str)
)
clean_sales["year"] = clean_sales["invoice_date"].dt.year
clean_sales["month"] = clean_sales["invoice_date"].dt.month
clean_sales["weekday"] = clean_sales["invoice_date"].dt.day_name()
clean_sales["hour"] = clean_sales["invoice_date"].dt.hour

# 调整字段顺序。
clean_sales = clean_sales[
    [
        "invoice_no",
        "stock_code",
        "description",
        "quantity",
        "unit_price",
        "sales_amount",
        "invoice_date",
        "order_date",
        "year_month",
        "year",
        "month",
        "weekday",
        "hour",
        "customer_id",
        "country",
    ]
]

clean_file = output_dir / "clean_sales.csv"
excluded_file = output_dir / "excluded_records.csv"
report_file = output_dir / "data_cleaning_report.txt"

clean_sales.to_csv(clean_file, index=False, encoding="utf-8-sig")
excluded_records.to_csv(excluded_file, index=False, encoding="utf-8-sig")

report_lines = [
    "电商订单数据清洗报告",
    "=" * 40,
    f"原始数据行数：{original_rows}",
    f"完全重复行数：{duplicate_rows}",
    f"去重后行数：{len(df)}",
    f"正常销售明细行数：{len(clean_sales)}",
    f"排除记录行数：{len(excluded_records)}",
    f"正常销售订单数：{clean_sales['invoice_no'].nunique()}",
    f"正常销售商品数：{clean_sales['stock_code'].nunique()}",
    f"正常销售客户数：{clean_sales['customer_id'].nunique()}",
    f"缺少客户编号的正常销售行数：{clean_sales['customer_id'].isna().sum()}",
    f"正常销售总金额：{clean_sales['sales_amount'].sum():.2f}",
]

report_file.write_text("\n".join(report_lines), encoding="utf-8")

print("\n".join(report_lines))
print("\n生成文件：")
print(clean_file.name)
print(excluded_file.name)
print(report_file.name)
print("\n数据清洗完成。")
