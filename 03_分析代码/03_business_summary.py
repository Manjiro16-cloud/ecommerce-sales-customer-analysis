from pathlib import Path

import pandas as pd


project_dir = Path(__file__).resolve().parent.parent
data_dir = project_dir / "02_清洗数据"
dashboard_dir = project_dir / "04_Excel经营看板"
report_dir = project_dir / "05_项目分析报告"

dashboard_dir.mkdir(exist_ok=True)
report_dir.mkdir(exist_ok=True)

input_file = data_dir / "clean_sales.csv"

print("正在读取清洗后的销售数据，请稍候...")
df = pd.read_csv(input_file, encoding="utf-8-sig")

df["invoice_date"] = pd.to_datetime(df["invoice_date"])
df["customer_id"] = pd.to_numeric(df["customer_id"], errors="coerce")

# 1. 总体 KPI
kpi = pd.DataFrame(
    {
        "指标": [
            "销售额",
            "销售明细行数",
            "订单数",
            "商品数",
            "客户数",
            "国家或地区数",
            "平均订单金额",
            "平均每单商品数量",
        ],
        "数值": [
            df["sales_amount"].sum(),
            len(df),
            df["invoice_no"].nunique(),
            df["stock_code"].nunique(),
            df["customer_id"].nunique(),
            df["country"].nunique(),
            df.groupby("invoice_no")["sales_amount"].sum().mean(),
            df.groupby("invoice_no")["quantity"].sum().mean(),
        ],
    }
)

# 2. 月度经营趋势
monthly = (
    df.groupby("year_month", as_index=False)
    .agg(
        sales_amount=("sales_amount", "sum"),
        order_count=("invoice_no", "nunique"),
        customer_count=("customer_id", "nunique"),
        item_quantity=("quantity", "sum"),
    )
    .sort_values("year_month")
)

monthly["average_order_value"] = (
    monthly["sales_amount"] / monthly["order_count"]
).round(2)

# 3. 国家或地区表现
country = (
    df.groupby("country", as_index=False)
    .agg(
        sales_amount=("sales_amount", "sum"),
        order_count=("invoice_no", "nunique"),
        customer_count=("customer_id", "nunique"),
        item_quantity=("quantity", "sum"),
    )
    .sort_values("sales_amount", ascending=False)
)

country["average_order_value"] = (
    country["sales_amount"] / country["order_count"]
).round(2)

# 4. 商品表现
# 排除人工调整、邮费、手续费、广告费等非普通商品交易项。
non_product_codes = {
    "M",
    "POST",
    "POSTAGE",
    "BANK CHARGES",
    "AMAZONFEE",
    "SAMPLES",
    "ADJUST",
    "ADJUST2",
    "DOT",
    "C2",
    "CRUK",
    "PADS",
    "DCGSSBOY",
    "DCGSSGIRL",
}

product_df = df[
    ~df["stock_code"].astype(str).str.upper().isin(non_product_codes)
].copy()

product = (
    product_df.groupby(["stock_code", "description"], as_index=False)
    .agg(
        sales_amount=("sales_amount", "sum"),
        order_count=("invoice_no", "nunique"),
        item_quantity=("quantity", "sum"),
    )
    .sort_values("sales_amount", ascending=False)
)

product["average_selling_price"] = (
    product["sales_amount"] / product["item_quantity"]
).round(2)

# 5. 星期与小时分布
weekday_order = [
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
    "Sunday",
]

weekday = (
    df.groupby("weekday", as_index=False)
    .agg(
        sales_amount=("sales_amount", "sum"),
        order_count=("invoice_no", "nunique"),
        item_quantity=("quantity", "sum"),
    )
)

weekday["weekday"] = pd.Categorical(
    weekday["weekday"],
    categories=weekday_order,
    ordered=True,
)
weekday = weekday.sort_values("weekday")

hour = (
    df.groupby("hour", as_index=False)
    .agg(
        sales_amount=("sales_amount", "sum"),
        order_count=("invoice_no", "nunique"),
        item_quantity=("quantity", "sum"),
    )
    .sort_values("hour")
)

# 6. 客户分层基础数据
customer_sales = (
    df.dropna(subset=["customer_id"])
    .groupby("customer_id", as_index=False)
    .agg(
        customer_sales=("sales_amount", "sum"),
        order_count=("invoice_no", "nunique"),
        item_quantity=("quantity", "sum"),
        last_purchase_date=("invoice_date", "max"),
    )
)

analysis_date = df["invoice_date"].max() + pd.Timedelta(days=1)

customer_sales["recency_days"] = (
    analysis_date - customer_sales["last_purchase_date"]
).dt.days

customer_sales["average_order_value"] = (
    customer_sales["customer_sales"] / customer_sales["order_count"]
).round(2)

customer_sales["customer_segment"] = "普通客户"

customer_sales.loc[
    (customer_sales["customer_sales"] >= customer_sales["customer_sales"].quantile(0.75))
    & (customer_sales["order_count"] >= customer_sales["order_count"].quantile(0.75)),
    "customer_segment",
] = "高价值客户"

customer_sales.loc[
    (customer_sales["recency_days"] <= customer_sales["recency_days"].quantile(0.25))
    & (customer_sales["order_count"] >= customer_sales["order_count"].median()),
    "customer_segment",
] = "活跃客户"

customer_sales.loc[
    customer_sales["recency_days"] >= customer_sales["recency_days"].quantile(0.75),
    "customer_segment",
] = "沉睡客户"

segment_summary = (
    customer_sales.groupby("customer_segment", as_index=False)
    .agg(
        customer_count=("customer_id", "count"),
        total_sales=("customer_sales", "sum"),
        average_customer_sales=("customer_sales", "mean"),
        average_order_count=("order_count", "mean"),
    )
    .sort_values("total_sales", ascending=False)
)

# 写入一个多工作表 Excel 文件。
output_file = dashboard_dir / "经营分析看板_重新生成.xlsx"

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    kpi.to_excel(writer, sheet_name="KPI总览", index=False)
    monthly.to_excel(writer, sheet_name="月度趋势", index=False)
    country.to_excel(writer, sheet_name="国家表现", index=False)
    product.head(100).to_excel(writer, sheet_name="商品TOP100", index=False)
    weekday.to_excel(writer, sheet_name="星期分布", index=False)
    hour.to_excel(writer, sheet_name="小时分布", index=False)
    customer_sales.to_excel(writer, sheet_name="客户明细", index=False)
    segment_summary.to_excel(writer, sheet_name="客户分层", index=False)

# 输出一份可直接阅读的项目初步结论。
top_country = country.iloc[0]
top_product = product.iloc[0]
top_month = monthly.loc[monthly["sales_amount"].idxmax()]
top_segment = segment_summary.iloc[0]

report_text = f"""
电商销售与客户经营分析：初步结论
================================

一、总体表现
销售额：{df["sales_amount"].sum():,.2f}
订单数：{df["invoice_no"].nunique():,}
普通商品数：{product["stock_code"].nunique():,}
客户数：{df["customer_id"].nunique():,}

二、关键发现
1. 销售额最高的国家或地区是：{top_country["country"]}
   销售额：{top_country["sales_amount"]:,.2f}
   订单数：{top_country["order_count"]:,}

2. 销售额最高的月份是：{top_month["year_month"]}
   销售额：{top_month["sales_amount"]:,.2f}
   订单数：{top_month["order_count"]:,}

3. 销售额最高的商品是：{top_product["description"]}
   商品编码：{top_product["stock_code"]}
   销售额：{top_product["sales_amount"]:,.2f}
   销量：{top_product["item_quantity"]:,}

4. 销售额贡献最高的客户分层是：{top_segment["customer_segment"]}
   客户数：{top_segment["customer_count"]:,}
   销售额：{top_segment["total_sales"]:,.2f}

三、数据限制
1. 部分销售明细缺少客户编号，因此客户分层不能覆盖所有销售记录。
2. 数据集没有库存字段，因此本项目不直接计算库存周转率。
3. 本项目使用公开数据集，不代表真实企业项目或本人实习经历。
"""

report_file = report_dir / "项目初步分析结论.txt"
report_file.write_text(report_text.strip(), encoding="utf-8")

print("\n分析完成。")
print("生成文件：", output_file)
print("生成报告：", report_file)
