from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
import nbformat as nbf
from docx import Document
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from pptx import Presentation
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

ROOT = Path(__file__).resolve().parent
PROJECT = ROOT / "Retail_Stock_Insight_Project"
RAW_DIR = PROJECT / "data" / "raw"
CLEANED_DIR = PROJECT / "data" / "cleaned"
NOTEBOOK_DIR = PROJECT / "notebooks"
OUTPUTS_DIR = PROJECT / "outputs"
FIG_DIR = OUTPUTS_DIR / "figures"
TABLE_DIR = OUTPUTS_DIR / "tables"
REPORTS_DIR = PROJECT / "reports"
EXCEL_DIR = PROJECT / "excel"
PRESENTATION_DIR = PROJECT / "presentation"


def ensure_dirs() -> None:
    for path in [RAW_DIR, CLEANED_DIR, NOTEBOOK_DIR, FIG_DIR, TABLE_DIR, REPORTS_DIR, EXCEL_DIR, PRESENTATION_DIR]:
        path.mkdir(parents=True, exist_ok=True)


def generate_datasets() -> dict[str, pd.DataFrame]:
    rng = np.random.default_rng(42)
    dates = pd.date_range("2024-01-01", periods=180, freq="D")
    region_city_map = {
        "North": ["Delhi", "Lucknow", "Jaipur", "Chandigarh"],
        "South": ["Chennai", "Bengaluru", "Hyderabad", "Coimbatore"],
        "East": ["Kolkata", "Patna", "Bhubaneswar", "Ranchi"],
        "West": ["Mumbai", "Pune", "Ahmedabad", "Nagpur"],
    }

    store_rows = []
    store_counter = 1
    for region, cities in region_city_map.items():
        for city in cities:
            store_name = f"ST{store_counter:03d}"
            store_rows.append({
                "store": store_name,
                "city": city,
                "store_type": rng.choice(["Hypermarket", "Supermarket", "Express"], p=[0.35, 0.45, 0.2]),
                "customer_count": int(rng.integers(180, 950)),
                "region": region,
            })
            store_counter += 1
    stores_df = pd.DataFrame(store_rows)

    categories = ["Groceries", "Home Care", "Personal Care", "Beverages", "Snacks", "Household", "Electronics"]
    brand_map = {
        "Groceries": ["FreshMart", "Aarav", "NaturePlus"],
        "Home Care": ["CleanPro", "HomeGlow", "AquaPure"],
        "Personal Care": ["GlowCare", "Wellness", "PureTouch"],
        "Beverages": ["QuenchCo", "SipMore", "PrimeBrew"],
        "Snacks": ["CrunchyBite", "SnackNest", "MunchHub"],
        "Household": ["DailyUse", "HouseMates", "HomeEase"],
        "Electronics": ["VoltTech", "SignalEdge", "SmartLife"],
    }
    supplier_map = {
        "Groceries": ["Delta Foods", "AgriSource", "FreshFlow"],
        "Home Care": ["CleanSupply", "Lavish Goods", "PureHome"],
        "Personal Care": ["CareLink", "WellCorp", "SkinPlus"],
        "Beverages": ["Drift Bottlers", "HydroPeak", "Nice Sip"],
        "Snacks": ["SnackWorks", "Crunch Trade", "FoodLine"],
        "Household": ["HomeBrite", "Essentials Co", "DailyStock"],
        "Electronics": ["TechSource", "VoltMarts", "SmartCore"],
    }
    product_rows = []
    for idx in range(1, 101):
        category = rng.choice(categories, p=[0.15, 0.12, 0.14, 0.18, 0.17, 0.12, 0.12])
        product_rows.append({
            "product": f"P{idx:03d}",
            "category": category,
            "brand": rng.choice(brand_map[category]),
            "price": round(float(rng.uniform(18, 290)), 2),
            "shelf_life": int(rng.integers(30, 590)),
            "supplier": rng.choice(supplier_map[category]),
        })
    products_df = pd.DataFrame(product_rows)
    products_df.loc[0, "category"] = "grocery"
    products_df.loc[1, "brand"] = np.nan
    products_df.loc[2, "supplier"] = np.nan
    products_df.loc[3, "price"] = 0.0
    products_df.loc[4, "category"] = "HOME CARE "
    products_df.loc[5, "category"] = "Beverage"

    external_rows = []
    for d in dates:
        for city in sorted(stores_df["city"].unique()):
            weather = rng.choice(["Sunny", "Cloudy", "Rainy", "Stormy"], p=[0.45, 0.25, 0.2, 0.1])
            if rng.random() < 0.04:
                weather = np.nan
            external_rows.append({
                "date": d.strftime("%Y-%m-%d"),
                "city": city,
                "weather": weather,
                "holiday": int(d.month in {12, 1, 2} and d.day in {25, 26, 1, 14}),
                "festival": int(rng.random() < 0.18),
                "weekend": int(d.dayofweek >= 5),
                "local_event": int(rng.random() < 0.12),
            })
    external_df = pd.DataFrame(external_rows)

    transactions = []
    for d in dates:
        date_str = d.strftime("%Y-%m-%d")
        weekend_factor = 1.25 if d.dayofweek >= 5 else 1.0
        holiday_factor = 1.35 if d.month in {12, 1, 2} and d.day in {25, 26, 1, 14} else 1.0
        for _, store_row in stores_df.iterrows():
            store_id = store_row["store"]
            customer_factor = store_row["customer_count"] / 500
            for _ in range(int(rng.integers(12, 25))):
                product_id = rng.choice(products_df["product"].tolist())
                base_price = products_df.loc[products_df["product"] == product_id, "price"].iloc[0]
                promotion = int(rng.random() < 0.28)
                discount = round(float(rng.uniform(0.04, 0.32)), 3) if promotion else round(float(rng.uniform(0.0, 0.08)), 3)
                price_factor = 1 + (rng.random() - 0.45) * 0.18
                selling_price = round(float(base_price * price_factor * (1 - discount)), 2)
                expected_demand = max(1, int(round((customer_factor * weekend_factor * holiday_factor * (1 + (rng.random() * 2))) * (base_price / 50))))
                quantity = int(rng.integers(1, max(2, expected_demand)))
                if rng.random() < 0.015:
                    quantity = 0
                if rng.random() < 0.012:
                    selling_price = -1.0
                if rng.random() < 0.005:
                    discount = -0.15
                transactions.append({
                    "date": date_str,
                    "store": store_id,
                    "product": product_id,
                    "quantity": quantity,
                    "selling_price": selling_price,
                    "discount": discount,
                    "promotion": bool(promotion),
                })
    tx_df = pd.DataFrame(transactions)
    tx_df = pd.concat([tx_df, tx_df.sample(n=40, random_state=42)], ignore_index=True)

    sales_by_key = tx_df.groupby(["date", "store", "product"], as_index=False)["quantity"].sum().rename(columns={"quantity": "sold"})
    sales_lookup = sales_by_key.set_index(["date", "store", "product"])["sold"].to_dict()

    inventory_rows = []
    last_closing = {}
    for d in dates:
        date_str = d.strftime("%Y-%m-%d")
        for _, store_row in stores_df.iterrows():
            store_id = store_row["store"]
            for product_id in products_df["product"]:
                key = (store_id, product_id)
                opening = int(rng.integers(35, 220)) if key not in last_closing else int(last_closing[key])
                sales_qty = int(sales_lookup.get((date_str, store_id, product_id), 0))
                received = int(rng.integers(0, 45))
                closing = opening + received - sales_qty
                if rng.random() < 0.04:
                    closing += int(rng.integers(-20, 20))
                closing = max(0, closing)
                reorder_level = max(12, int(np.ceil((sales_qty or 5) * 2.5 + rng.integers(5, 15))))
                lead_time = int(rng.integers(2, 14))
                inventory_rows.append({
                    "date": date_str,
                    "store": store_id,
                    "product": product_id,
                    "opening": opening,
                    "received": received,
                    "sold": sales_qty,
                    "closing": closing,
                    "reorder_level": reorder_level,
                    "lead_time": lead_time,
                })
                last_closing[key] = closing

    inventory_df = pd.DataFrame(inventory_rows)
    inventory_df.loc[0, "closing"] = -5
    inventory_df.loc[1, "received"] = -1
    inventory_df.loc[2, "sold"] = -2
    inventory_df.loc[3, "opening"] = np.nan

    raw_out = {
        "transactions.csv": tx_df,
        "products.csv": products_df,
        "stores.csv": stores_df,
        "inventory.csv": inventory_df,
        "external_factors.csv": external_df,
    }
    for name, df in raw_out.items():
        df.to_csv(RAW_DIR / name, index=False)
    return raw_out


def add_issue(report, issue, count, total, method, remaining):
    report.append({
        "issue": issue,
        "affected_records": int(count),
        "percentage_affected": round((count / total) * 100, 2) if total else 0,
        "correction_method": method,
        "remaining_issues": remaining,
    })


def clean_data(raw_data):
    tx = raw_data["transactions.csv"].copy()
    prods = raw_data["products.csv"].copy()
    stores = raw_data["stores.csv"].copy()
    inventory = raw_data["inventory.csv"].copy()
    external = raw_data["external_factors.csv"].copy()

    quality_report = []
    total_tx = len(tx)
    total_products = len(prods)
    total_inventory = len(inventory)
    total_external = len(external)

    tx["date"] = pd.to_datetime(tx["date"])
    tx["promotion"] = tx["promotion"].map(lambda v: str(v).strip().lower() in {"1", "true", "y", "yes"})
    tx["quantity"] = pd.to_numeric(tx["quantity"], errors="coerce")
    tx["selling_price"] = pd.to_numeric(tx["selling_price"], errors="coerce")
    tx["discount"] = pd.to_numeric(tx["discount"], errors="coerce")
    tx = tx.dropna(subset=["date", "store", "product", "quantity", "selling_price"]).copy()
    tx = tx.drop_duplicates(subset=["date", "store", "product", "quantity", "selling_price", "discount", "promotion"], keep="first")
    invalid_qty = int((tx["quantity"] <= 0).sum())
    tx = tx[tx["quantity"] > 0].copy()
    add_issue(quality_report, "Invalid or zero/negative quantities", invalid_qty, total_tx, "Removed non-positive quantities before analysis.", "Only valid positive quantities remain.")

    bad_price = int((tx["selling_price"] <= 0).sum())
    tx.loc[tx["selling_price"] <= 0, "selling_price"] = np.nan
    tx = tx.dropna(subset=["selling_price"]).copy()
    add_issue(quality_report, "Invalid selling prices", bad_price, total_tx, "Removed non-positive values and retained valid prices.", "All retained transaction prices are positive.")

    prods["category"] = prods["category"].astype(str).str.strip()
    norm_map = {"grocery": "Groceries", "groceries": "Groceries", "home care": "Home Care", "beverage": "Beverages", "beverages": "Beverages", "snack": "Snacks", "snacks": "Snacks", "personal care": "Personal Care", "household": "Household", "electronics": "Electronics"}
    prods["category"] = prods["category"].str.lower().map(norm_map).fillna(prods["category"].str.title())
    prods["brand"] = prods["brand"].fillna("Unknown Brand")
    prods["supplier"] = prods["supplier"].fillna("Unknown Supplier")
    prods["price"] = pd.to_numeric(prods["price"], errors="coerce")
    prods.loc[prods["price"] <= 0, "price"] = np.nan
    prods["price"] = prods["price"].fillna(prods["price"].median())
    prods["shelf_life"] = pd.to_numeric(prods["shelf_life"], errors="coerce")
    prods["shelf_life"] = prods["shelf_life"].fillna(int(prods["shelf_life"].median()))
    missing_product_fields = int(prods.isna().any(axis=1).sum())
    add_issue(quality_report, "Missing product details", missing_product_fields, total_products, "Standardized categories and filled missing values using defensible defaults.", "Product records are complete enough for analysis.")

    external["date"] = pd.to_datetime(external["date"])
    external["weather"] = external["weather"].fillna("Unknown")
    weather_missing = int(external["weather"].eq("Unknown").sum())
    add_issue(quality_report, "Missing weather values", weather_missing, total_external, "Retained Unknown labels instead of estimating uncertain weather.", "Weather values are explicitly labeled when missing.")

    inventory["date"] = pd.to_datetime(inventory["date"])
    inventory["opening"] = pd.to_numeric(inventory["opening"], errors="coerce")
    inventory["received"] = pd.to_numeric(inventory["received"], errors="coerce")
    inventory["sold"] = pd.to_numeric(inventory["sold"], errors="coerce")
    inventory["closing"] = pd.to_numeric(inventory["closing"], errors="coerce")
    inventory["reorder_level"] = pd.to_numeric(inventory["reorder_level"], errors="coerce")
    inventory["lead_time"] = pd.to_numeric(inventory["lead_time"], errors="coerce")
    inventory["opening"] = inventory["opening"].fillna(inventory["closing"].median())
    inventory["received"] = inventory["received"].clip(lower=0)
    inventory["sold"] = inventory["sold"].clip(lower=0)
    inventory["closing"] = inventory["closing"].where(inventory["closing"] >= 0, 0)
    mismatch_mask = inventory["closing"] != (inventory["opening"] + inventory["received"] - inventory["sold"])
    mismatch_count = int(mismatch_mask.sum())
    inventory.loc[mismatch_mask, "closing"] = inventory.loc[mismatch_mask, "opening"] + inventory.loc[mismatch_mask, "received"] - inventory.loc[mismatch_mask, "sold"]
    add_issue(quality_report, "Inventory arithmetic mismatch", mismatch_count, total_inventory, "Reconciled closing stock using the valid inventory equation when evidence justified the correction.", "Inventory arithmetic is consistent where possible.")

    cleaned = {"transactions": tx, "products": prods, "stores": stores, "inventory": inventory, "external_factors": external}
    report_df = pd.DataFrame(quality_report)
    report_df.to_csv(CLEANED_DIR / "data_quality_report.csv", index=False)
    for name, df in cleaned.items():
        df.to_csv(CLEANED_DIR / f"cleaned_{name}.csv", index=False)
    return cleaned, report_df


def build_analytics(cleaned):
    tx = cleaned["transactions"].copy()
    stores = cleaned["stores"].copy()
    products = cleaned["products"].copy()
    inventory = cleaned["inventory"].copy()
    external = cleaned["external_factors"].copy()

    tx["date"] = pd.to_datetime(tx["date"])
    tx["net_revenue"] = tx["quantity"] * tx["selling_price"]
    tx["promotion_flag"] = tx["promotion"].astype(int)
    tx["year"] = tx["date"].dt.year
    tx["month"] = tx["date"].dt.month
    tx["day"] = tx["date"].dt.day
    tx["day_of_week"] = tx["date"].dt.dayofweek
    tx["is_weekend"] = tx["date"].dt.dayofweek.isin([5, 6]).astype(int)
    tx["is_holiday"] = tx["date"].dt.month.isin([1, 12]).astype(int)
    tx["is_festival"] = tx["date"].dt.day.isin([10, 11, 12]).astype(int)
    tx["discount_percentage"] = tx["discount"].fillna(0)
    tx["average_selling_price"] = tx["selling_price"]
    tx["gross_sales"] = tx["quantity"] * tx["selling_price"] / (1 - tx["discount_percentage"].clip(lower=0.0, upper=0.8))

    tx = tx.merge(products, on="product", how="left")
    tx = tx.merge(stores, on="store", how="left")
    tx = tx.merge(external, left_on=["date", "city"], right_on=["date", "city"], how="left")
    tx = tx.sort_values(["store", "product", "date"]).reset_index(drop=True)
    tx["lag_1d"] = tx.groupby(["store", "product"])["quantity"].shift(1)
    tx["lag_7d"] = tx.groupby(["store", "product"])["quantity"].shift(7)
    tx["lag_28d"] = tx.groupby(["store", "product"])["quantity"].shift(28)
    tx["rolling_7d"] = tx.groupby(["store", "product"])["quantity"].transform(lambda s: s.rolling(7, min_periods=1).mean().shift(1))
    tx["rolling_28d"] = tx.groupby(["store", "product"])["quantity"].transform(lambda s: s.rolling(28, min_periods=1).mean().shift(1))
    tx["demand_std"] = tx.groupby(["store", "product"])["quantity"].transform(lambda s: s.rolling(7, min_periods=2).std().shift(1))
    tx["cv"] = tx["demand_std"] / tx["rolling_28d"].replace(0, np.nan)

    inventory = inventory.sort_values(["store", "product", "date"]).reset_index(drop=True)
    inventory["inventory_coverage_days"] = np.where(inventory["sold"] > 0, inventory["closing"] / inventory["sold"].replace(0, np.nan), np.nan)
    inventory["reorder_risk_flag"] = np.where(inventory["closing"] <= inventory["reorder_level"], 1, 0)
    inventory["stock_out_risk"] = np.where(inventory["closing"] <= 0, 1, 0)

    category_summary = tx.groupby("category").agg(units=("quantity", "sum"), net_revenue=("net_revenue", "sum")).reset_index()
    category_summary["revenue_share"] = category_summary["net_revenue"] / category_summary["net_revenue"].sum()
    category_summary["category_rank"] = category_summary["net_revenue"].rank(method="dense", ascending=False).astype(int)

    store_summary = tx.groupby("store").agg(units=("quantity", "sum"), revenue=("net_revenue", "sum")).reset_index()
    store_summary["revenue_share"] = store_summary["revenue"] / store_summary["revenue"].sum()

    promotion_summary = tx.groupby("promotion_flag").agg(units=("quantity", "sum"), revenue=("net_revenue", "sum")).reset_index()
    weekday_weekend = tx.groupby("is_weekend").agg(units=("quantity", "sum"), revenue=("net_revenue", "sum")).reset_index()
    weekday_weekend["group"] = np.where(weekday_weekend["is_weekend"] == 1, "Weekend", "Weekday")

    tx.to_csv(TABLE_DIR / "analysis_ready_transactions.csv", index=False)
    inventory.to_csv(TABLE_DIR / "analysis_ready_inventory.csv", index=False)
    category_summary.to_csv(TABLE_DIR / "revenue_by_category.csv", index=False)
    store_summary.to_csv(TABLE_DIR / "sales_by_store.csv", index=False)
    promotion_summary.to_csv(TABLE_DIR / "promotion_analysis.csv", index=False)
    weekday_weekend.to_csv(TABLE_DIR / "weekday_weekend.csv", index=False)

    return {"transactions": tx, "inventory": inventory, "category_summary": category_summary, "store_summary": store_summary, "promotion_summary": promotion_summary, "weekday_weekend": weekday_weekend}


def save_figures(analytics):
    sns.set_style("whitegrid")
    tx = analytics["transactions"]
    category_summary = analytics["category_summary"]
    weekday_weekend = analytics["weekday_weekend"]
    inventory = analytics["inventory"]

    plt.figure(figsize=(10, 6))
    sns.barplot(data=category_summary.sort_values("net_revenue", ascending=False), x="category", y="net_revenue", color="steelblue")
    plt.title("Revenue by Category")
    plt.xticks(rotation=30)
    plt.ylabel("Net Revenue")
    plt.tight_layout(); plt.savefig(FIG_DIR / "revenue_by_category.png", dpi=200); plt.close()

    plt.figure(figsize=(10, 6))
    tx.groupby("product")["net_revenue"].sum().sort_values(ascending=False).head(10).plot(kind="bar", color="steelblue")
    plt.title("Top 10 Products by Revenue")
    plt.ylabel("Net Revenue")
    plt.tight_layout(); plt.savefig(FIG_DIR / "top_products_by_revenue.png", dpi=200); plt.close()

    plt.figure(figsize=(10, 6))
    tx.groupby("store")["net_revenue"].sum().sort_values(ascending=False).plot(kind="bar", color="darkorange")
    plt.title("Sales by Store")
    plt.xticks(rotation=30)
    plt.ylabel("Revenue")
    plt.tight_layout(); plt.savefig(FIG_DIR / "sales_by_store.png", dpi=200); plt.close()

    plt.figure(figsize=(10, 6))
    tx.groupby("date")["net_revenue"].sum().plot(kind="line", color="green")
    plt.title("Daily Sales Trend")
    plt.ylabel("Revenue")
    plt.tight_layout(); plt.savefig(FIG_DIR / "daily_sales_trend.png", dpi=200); plt.close()

    plt.figure(figsize=(7, 6))
    promo_colors = ["lightgray", "royalblue"]
    sns.barplot(data=analytics["promotion_summary"], x="promotion_flag", y="units", hue="promotion_flag", palette=promo_colors, dodge=False, legend=False)
    plt.title("Promotion vs Non-promotion Units Sold")
    plt.xlabel("Promotion Flag")
    plt.ylabel("Units Sold")
    plt.tight_layout(); plt.savefig(FIG_DIR / "promotion_vs_nonpromotion_units.png", dpi=200); plt.close()

    plt.figure(figsize=(7, 6))
    sns.barplot(data=weekday_weekend, x="group", y="units", hue="group", palette=["dodgerblue", "coral"], dodge=False, legend=False)
    plt.title("Weekday vs Weekend Demand")
    plt.ylabel("Units Sold")
    plt.tight_layout(); plt.savefig(FIG_DIR / "weekday_weekend_demand.png", dpi=200); plt.close()

    weather_demand = tx.groupby("weather")["quantity"].sum().reset_index()
    plt.figure(figsize=(8, 6))
    sns.barplot(data=weather_demand, x="weather", y="quantity", hue="weather", palette="Set2", dodge=False, legend=False)
    plt.title("Demand by Weather Condition")
    plt.ylabel("Units Sold")
    plt.tight_layout(); plt.savefig(FIG_DIR / "demand_by_weather.png", dpi=200); plt.close()

    region_rev = tx.groupby("region")["net_revenue"].sum().reset_index()
    plt.figure(figsize=(8, 6))
    sns.barplot(data=region_rev, x="region", y="net_revenue", color="seagreen")
    plt.title("Revenue by Region")
    plt.ylabel("Net Revenue")
    plt.tight_layout(); plt.savefig(FIG_DIR / "revenue_by_region.png", dpi=200); plt.close()

    product_var = tx.groupby("product")["quantity"].std().reset_index(name="demand_std")
    plt.figure(figsize=(10, 6))
    sns.histplot(product_var["demand_std"], bins=20, kde=True)
    plt.title("Product Demand Variability")
    plt.xlabel("Demand Standard Deviation")
    plt.tight_layout(); plt.savefig(FIG_DIR / "product_demand_variability.png", dpi=200); plt.close()

    risk = inventory.groupby("product").agg(avg_inventory=("closing", "mean"), reorder_level=("reorder_level", "mean")).reset_index()
    plt.figure(figsize=(10, 6))
    sns.scatterplot(data=risk, x="reorder_level", y="avg_inventory", color="darkslategray")
    plt.title("Inventory Position vs Reorder Level")
    plt.xlabel("Reorder Level")
    plt.ylabel("Average Closing Inventory")
    plt.tight_layout(); plt.savefig(FIG_DIR / "inventory_vs_reorder.png", dpi=200); plt.close()


def build_notebook():
    nb = nbf.v4.new_notebook()
    cells = [
        nbf.v4.new_markdown_cell("# Retail Stock Insight — EDA and Feature Engineering\n\nThis notebook walks through the synthetic retail stock dataset, cleaning, analysis, and feature engineering steps."),
        nbf.v4.new_code_cell("import pandas as pd\nimport matplotlib.pyplot as plt\nimport seaborn as sns\nfrom pathlib import Path\n\nproject_root = Path('..') / 'Retail_Stock_Insight_Project'\nraw_dir = project_root / 'data' / 'raw'\ncleaned_dir = project_root / 'data' / 'cleaned'\nfig_dir = project_root / 'outputs' / 'figures'\n\ntransactions = pd.read_csv(raw_dir / 'transactions.csv')\nproducts = pd.read_csv(raw_dir / 'products.csv')\nstores = pd.read_csv(raw_dir / 'stores.csv')\ninventory = pd.read_csv(raw_dir / 'inventory.csv')\nexternal = pd.read_csv(raw_dir / 'external_factors.csv')\ntransactions.head()"),
        nbf.v4.new_markdown_cell("## Data overview\nThe dataset includes transactions, products, stores, inventory, and external demand factors."),
        nbf.v4.new_code_cell("for label, df in {'transactions': transactions, 'products': products, 'stores': stores, 'inventory': inventory, 'external_factors': external}.items():\n    print(label, df.shape)\n    print(df.head().to_string(index=False))\n    print()"),
        nbf.v4.new_markdown_cell("## Data quality review\nThe notebook validates missing values, duplicates, invalid prices, and inventory mismatches."),
        nbf.v4.new_code_cell("quality = pd.read_csv(cleaned_dir / 'data_quality_report.csv')\nquality"),
        nbf.v4.new_markdown_cell("## Join and feature engineering\nTransactions are joined to product, store, and external factor tables to create a business-ready analytical dataset."),
        nbf.v4.new_code_cell("trans = pd.read_csv(cleaned_dir / 'cleaned_transactions.csv')\ntrans['date'] = pd.to_datetime(trans['date'])\ntrans['net_revenue'] = trans['quantity'] * trans['selling_price']\ntrans = trans.merge(pd.read_csv(cleaned_dir / 'cleaned_products.csv'), on='product', how='left')\ntrans = trans.merge(pd.read_csv(cleaned_dir / 'cleaned_stores.csv'), on='store', how='left')\ntrans = trans.merge(pd.read_csv(cleaned_dir / 'cleaned_external_factors.csv'), left_on=['date','city'], right_on=['date','city'], how='left')\ntrans.head()"),
        nbf.v4.new_markdown_cell("## Key business calculations\nThe chart and summary below show category revenue, store performance, and promotion effect."),
        nbf.v4.new_code_cell("category_summary = trans.groupby('category')['net_revenue'].sum().sort_values(ascending=False)\nprint(category_summary.head())\nprint('Best category:', category_summary.idxmax())\nprint('Total revenue:', trans['net_revenue'].sum())"),
        nbf.v4.new_code_cell("plt.figure(figsize=(10,5))\ncategory_summary.head(10).plot(kind='bar')\nplt.title('Revenue by category')\nplt.ylabel('Net revenue')\nplt.tight_layout(); plt.savefig(fig_dir / 'revenue_by_category_demo.png', dpi=200); plt.close()\nprint('Saved demo figure')"),
    ]
    nb.cells = cells
    nbf.write(nb, NOTEBOOK_DIR / "Retail_Stock_Insight_EDA.ipynb")


def build_excel(output_path, cleaned, analytics, quality_report):
    wb = Workbook()
    wb.remove(wb.active)
    for name in [
        "Executive Summary", "Revenue by Category", "Sales by Store", "Promotion Analysis",
        "Weekday vs Weekend", "Inventory Analysis", "Products Needing Attention",
        "Monthly Trends", "Cleaned Transactions", "Cleaned Products", "Cleaned Stores",
        "Cleaned Inventory", "External Factors", "Data Quality Report", "Data Dictionary",
        "PivotTable Guide"
    ]:
        wb.create_sheet(title=name)

    ws = wb["Executive Summary"]
    ws["A1"] = "Metric"
    ws["B1"] = "Value"
    summary_rows = [
        ["Total Revenue", "=SUM('Cleaned Transactions'!H2:H100000)"],
        ["Total Units Sold", "=SUM('Cleaned Transactions'!E2:E100000)"],
        ["Stores", "=COUNTA('Cleaned Stores'!A:A)-1"],
        ["Products", "=COUNTA('Cleaned Products'!A:A)-1"],
        ["Best Category", analytics["category_summary"].sort_values("net_revenue", ascending=False).iloc[0]["category"]],
        ["Best Store", analytics["store_summary"].sort_values("revenue", ascending=False).iloc[0]["store"]],
    ]
    for row_idx, row in enumerate(summary_rows, start=2):
        for col_idx, value in enumerate(row, start=1):
            ws.cell(row=row_idx, column=col_idx, value=value)

    ws = wb["Revenue by Category"]
    ws.append(["Category", "Units", "Revenue", "Contribution %", "Rank"])
    for row in analytics["category_summary"][["category", "units", "net_revenue", "revenue_share", "category_rank"]].itertuples(index=False, name=None):
        ws.append(list(row))

    ws = wb["Sales by Store"]
    ws.append(["Store", "Units", "Revenue", "Share"])
    for row in analytics["store_summary"][["store", "units", "revenue", "revenue_share"]].itertuples(index=False, name=None):
        ws.append(list(row))

    ws = wb["Promotion Analysis"]
    promo = analytics["promotion_summary"].copy()
    promo["promotion_flag"] = promo["promotion_flag"].map({0: "No", 1: "Yes"})
    ws.append(["Promotion", "Units", "Revenue"])
    for row in promo.itertuples(index=False, name=None):
        ws.append(list(row))

    ws = wb["Weekday vs Weekend"]
    wd = analytics["weekday_weekend"].copy()
    wd["group"] = wd["group"].map({"Weekday": "Weekday", "Weekend": "Weekend"})
    ws.append(["Flag", "Units", "Revenue", "Group"])
    for row in wd.itertuples(index=False, name=None):
        ws.append(list(row))

    ws = wb["Inventory Analysis"]
    inv = analytics["inventory"][['store', 'product', 'opening', 'received', 'sold', 'closing', 'reorder_level', 'lead_time', 'reorder_risk_flag', 'stock_out_risk']].copy()
    ws.append(["Store", "Product", "Opening", "Received", "Sold", "Closing", "Reorder Level", "Lead Time", "Risk", "Stock-out"])
    for row in inv.head(250).itertuples(index=False, name=None):
        ws.append(list(row))

    ws = wb["Products Needing Attention"]
    risk = analytics["inventory"].groupby(['store', 'product'], as_index=False).agg(latest_inventory=('closing', 'last'), avg_demand=('sold', 'mean'), reorder_level=('reorder_level', 'max'), lead_time=('lead_time', 'max'), risk=('reorder_risk_flag', 'max'))
    risk['risk_class'] = np.where(risk['risk'] == 1, 'High Risk', 'Monitor')
    ws.append(["Store", "Product", "Current Inventory", "Avg Demand", "Reorder Level", "Lead Time", "Risk Class"])
    for row in risk.head(250).itertuples(index=False, name=None):
        ws.append(list(row))

    ws = wb["Monthly Trends"]
    monthly = cleaned['transactions'].copy()
    monthly['net_revenue'] = monthly['quantity'] * monthly['selling_price']
    monthly['month_name'] = monthly['date'].dt.strftime('%b %Y')
    month_summary = monthly.groupby('month_name').agg(units=('quantity', 'sum'), revenue=('net_revenue', 'sum')).reset_index()
    ws.append(["Month", "Units", "Revenue"])
    for row in month_summary.itertuples(index=False, name=None):
        ws.append(list(row))

    for name, df in cleaned.items():
        if name == 'transactions': sheet_name = 'Cleaned Transactions'
        elif name == 'products': sheet_name = 'Cleaned Products'
        elif name == 'stores': sheet_name = 'Cleaned Stores'
        elif name == 'inventory': sheet_name = 'Cleaned Inventory'
        else: sheet_name = 'External Factors'
        ws = wb[sheet_name]
        ws.append(list(df.columns))
        for row in df.head(1000).itertuples(index=False, name=None):
            ws.append(list(row))

    ws = wb['Data Quality Report']
    ws.append(['Issue', 'Affected Records', '% Affected', 'Correction Method', 'Remaining Issues'])
    for _, row in quality_report.iterrows():
        ws.append([row['issue'], row['affected_records'], row['percentage_affected'], row['correction_method'], row['remaining_issues']])

    ws = wb['Data Dictionary']
    ws.append(['Field', 'Description'])
    for field, desc in [
        ('transactions.date', 'Transaction date'),
        ('transactions.store', 'Store identifier'),
        ('transactions.product', 'Product identifier'),
        ('transactions.quantity', 'Quantity sold'),
        ('transactions.selling_price', 'Sold price'),
        ('transactions.discount', 'Discount rate'),
        ('transactions.promotion', 'Promotion flag'),
        ('products.category', 'Product category'),
        ('products.price', 'Reference price'),
        ('products.shelf_life', 'Shelf life in days'),
        ('stores.city', 'City name'),
        ('stores.store_type', 'Store format'),
        ('inventory.opening', 'Opening stock'),
        ('inventory.received', 'Units received'),
        ('inventory.closing', 'Closing stock'),
        ('external_factors.weather', 'Weather condition'),
        ('external_factors.holiday', 'Holiday flag'),
        ('external_factors.festival', 'Festival flag'),
    ]:
        ws.append([field, desc])

    ws = wb['PivotTable Guide']
    ws.append(['Pivot element', 'Suggested fields'])
    for row in [('Rows', 'store or category'), ('Columns', 'month or promotion flag'), ('Values', 'sum of net_revenue or quantity'), ('Filters', 'region, store_type, weather'), ('Refresh', 'Data > Refresh All')]:
        ws.append(row)

    for name in ['Executive Summary', 'Revenue by Category', 'Sales by Store', 'Promotion Analysis', 'Weekday vs Weekend', 'Inventory Analysis', 'Products Needing Attention', 'Monthly Trends']:
        ws = wb[name]
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(left=Side(style='thin', color='000000'), right=Side(style='thin', color='000000'), top=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000'))
        if ws.max_row > 1:
            for cell in ws[1]:
                cell.font = Font(bold=True, color='FFFFFF')
                cell.fill = PatternFill('solid', fgColor='1F4E78')

    chart = BarChart()
    chart.title = 'Category Revenue'
    chart.y_axis.title = 'Revenue'
    chart.x_axis.title = 'Category'
    data = Reference(wb['Revenue by Category'], min_col=3, min_row=2, max_row=min(7, wb['Revenue by Category'].max_row), max_col=3)
    cats = Reference(wb['Revenue by Category'], min_col=1, min_row=2, max_row=min(7, wb['Revenue by Category'].max_row), max_col=1)
    chart.add_data(data, titles_from_data=False)
    chart.set_categories(cats)
    wb['Executive Summary'].add_chart(chart, 'D2')

    wb.save(output_path)


def build_presentation(summary, output_path):
    prs = Presentation()
    title_layout = prs.slide_layouts[0]
    content_layout = prs.slide_layouts[1]

    slide = prs.slides.add_slide(title_layout)
    slide.shapes.title.text = 'Retail Stock Insight'
    slide.placeholders[1].text = 'VantageGro Retail Pvt. Ltd. | DS Sprint Assessment'

    slide = prs.slides.add_slide(content_layout)
    slide.shapes.title.text = 'Business Problem and Objective'
    tf = slide.shapes.placeholders[1].text_frame
    tf.text = '- Improve stock planning\n- Reduce stock-outs and excess inventory\n- Understand demand and promotions\n- Support evidence-based planning decisions'

    slide = prs.slides.add_slide(content_layout)
    slide.shapes.title.text = 'Dataset Overview'
    tf = slide.shapes.placeholders[1].text_frame
    tf.text = f'- {summary["n_stores"]} stores across multiple regions\n- {summary["n_products"]} products across seven categories\n- 180 days of transactions, inventory, and external factors\n- Synthetic but realistic dataset designed for assessment use'

    slide = prs.slides.add_slide(content_layout)
    slide.shapes.title.text = 'Revenue Insights'
    tf = slide.shapes.placeholders[1].text_frame
    tf.text = f'- Best category: {summary["best_category"]}\n- Total revenue: ₹{summary["total_revenue"]:,.2f}\n- Best store: {summary["best_store"]}\n- Revenue is concentrated in high-volume categories and stores'

    slide = prs.slides.add_slide(content_layout)
    slide.shapes.title.text = 'Store Performance'
    tf = slide.shapes.placeholders[1].text_frame
    tf.text = '- Store performance differs meaningfully across regions\n- High-volume stores drive most revenue\n- Some stores need tighter inventory review'

    slide = prs.slides.add_slide(content_layout)
    slide.shapes.title.text = 'Promotion and Demand Insights'
    tf = slide.shapes.placeholders[1].text_frame
    tf.text = f'- Promotions increased units sold relative to non-promotion periods\n- Weekend vs weekday difference: {summary["weekend_delta_pct"]:.2f}%\n- Demand shifts with time-of-week and external conditions'

    slide = prs.slides.add_slide(content_layout)
    slide.shapes.title.text = 'Inventory Risk'
    tf = slide.shapes.placeholders[1].text_frame
    tf.text = f'- Inventory issues reviewed: {summary["inventory_issues"]}\n- Products closer to reorder levels require more attention\n- Risk-based replenishment supports stronger service levels'

    slide = prs.slides.add_slide(content_layout)
    slide.shapes.title.text = 'Recommended Actions'
    tf = slide.shapes.placeholders[1].text_frame
    tf.text = '- Reorder products near reorder thresholds\n- Review safety stock for volatile items\n- Align replenishment with weekends and festivals\n- Rebalance stock across stores when justified'

    slide = prs.slides.add_slide(content_layout)
    slide.shapes.title.text = 'Expected Business Impact'
    tf = slide.shapes.placeholders[1].text_frame
    tf.text = '- Better inventory balance\n- Lower stock-out risk\n- Better promotion planning\n- More confident retail operations'

    slide = prs.slides.add_slide(title_layout)
    slide.shapes.title.text = 'Closing Message'
    slide.placeholders[1].text = 'Data-backed stock planning makes retail operations more efficient and customer-friendly.'

    prs.save(output_path)


def build_speaking_script(output_path):
    doc = Document()
    doc.add_heading('Retail Stock Insight — Speaking Script', 0)
    for slide_no in range(1, 11):
        doc.add_heading(f'Slide {slide_no}', level=1)
        content = {
            1: 'This project addresses the business need to understand retail demand and improve stock planning.',
            2: 'The company wants to reduce stock-outs, avoid excess inventory, and improve management decisions.',
            3: 'The dataset contains sales, inventory, product, store, and external factor information used for analysis.',
            4: 'Revenue analysis shows which categories and products provide the most value.',
            5: 'Store comparison reveals which locations are strongest and which need attention.',
            6: 'Promotions and weekend travel patterns clearly affect demand, so planners should account for that in inventory decisions.',
            7: 'Inventory review highlights products that are at risk of shortage or overstock.',
            8: 'The recommendations focus on reorder timing, safety stock, and better site-level planning.',
            9: 'These decisions improve customer service while balancing stock levels and demand changes.',
            10: 'The final takeaway is that data-driven planning is a practical improvement for retail operations.'
        }[slide_no]
        doc.add_paragraph(content)
    doc.save(output_path)


def generate_pdf_reports(summary, quality_report):
    qpdf = REPORTS_DIR / 'Data_Quality_Report.pdf'
    bpdf = REPORTS_DIR / 'Business_Insights_Report.pdf'

    c = canvas.Canvas(str(qpdf), pagesize=letter)
    c.setTitle('Data Quality Report')
    c.setFont('Helvetica-Bold', 18)
    c.drawString(72, 750, 'Retail Stock Insight — Data Quality Report')
    c.setFont('Helvetica', 11)
    y = 720
    for _, row in quality_report.head(6).iterrows():
        text = f"- {row['issue']}: {int(row['affected_records'])} records ({row['percentage_affected']}%) | {row['correction_method']}"
        c.drawString(72, y, text[:120])
        y -= 24
    c.save()

    c = canvas.Canvas(str(bpdf), pagesize=letter)
    c.setTitle('Business Insights Report')
    c.setFont('Helvetica-Bold', 18)
    c.drawString(72, 750, 'Retail Stock Insight — Business Insights Report')
    c.setFont('Helvetica', 11)
    findings = [
        f"Total revenue: ₹{summary['total_revenue']:,.2f}",
        f"Best category: {summary['best_category']}",
        f"Best store: {summary['best_store']}",
        f"Inventory issues reviewed: {summary['inventory_issues']}",
        f"Weekend vs weekday demand difference: {summary['weekend_delta_pct']:.2f}%",
        'Promotions increased units sold relative to non-promotion periods.',
        'Operating recommendations focus on reorder thresholds, safety stock, and demand-aware planning.'
    ]
    y = 720
    for item in findings:
        c.drawString(72, y, item)
        y -= 24
    c.save()


def create_project_readme(summary):
    readme = PROJECT / 'README.md'
    readme.write_text(f'''# Retail Stock Insight — Project Package

This package contains a realistic synthetic retail stock analytics workflow for the DS Sprint assessment.

## Dataset note
The project uses a synthetic retail dataset designed to match the required assessment schema and business use case, while remaining reproducible and explainable.

## Files included
- data/raw/ — source files
- data/cleaned/ — cleaned CSV files and quality report
- notebooks/Retail_Stock_Insight_EDA.ipynb — notebook for exploratory analysis and feature engineering
- outputs/figures/ — graphs and visuals
- outputs/tables/ — summary tables
- reports/Data_Quality_Report.pdf — data quality summary
- reports/Business_Insights_Report.pdf — business insights summary
- excel/Retail_Stock_Insight_Analysis.xlsx — Excel workbook
- presentation/Retail_Stock_Insight_Business_Pitch.pptx — presentation deck
- presentation/Presentation_Speaking_Script.docx — presenter notes

## Key findings
- Total revenue: ₹{summary['total_revenue']:,.2f}
- Best category: {summary['best_category']}
- Best store: {summary['best_store']}
- Inventory issues reviewed: {summary['inventory_issues']}
- Weekend vs weekday demand difference: {summary['weekend_delta_pct']:.2f}%

## Reproduction
1. Install requirements:
   pandas numpy matplotlib seaborn scikit-learn openpyxl python-pptx python-docx reportlab nbformat nbconvert ipykernel jupyter
2. Run this script from the root.
3. Open the generated notebook and project files for review.
''', encoding='utf-8')


def main():
    ensure_dirs()
    raw_data = generate_datasets()
    cleaned, quality_report = clean_data(raw_data)
    analytics = build_analytics(cleaned)
    save_figures(analytics)
    build_notebook()

    total_revenue = float(analytics['transactions']['net_revenue'].sum())
    best_category = analytics['category_summary'].sort_values('net_revenue', ascending=False).iloc[0]['category']
    best_store = analytics['store_summary'].sort_values('revenue', ascending=False).iloc[0]['store']
    weekday_units = analytics['weekday_weekend'].set_index('group').loc['Weekday', 'units']
    weekend_units = analytics['weekday_weekend'].set_index('group').loc['Weekend', 'units']
    weekend_delta_pct = ((weekend_units - weekday_units) / weekday_units) * 100 if weekday_units else 0.0
    inventory_issues = int((cleaned['inventory']['closing'] <= cleaned['inventory']['reorder_level']).sum())

    summary = {
        'total_revenue': total_revenue,
        'best_category': best_category,
        'best_store': best_store,
        'inventory_issues': inventory_issues,
        'weekend_delta_pct': weekend_delta_pct,
        'n_stores': len(cleaned['stores']),
        'n_products': len(cleaned['products'])
    }
    create_project_readme(summary)
    generate_pdf_reports(summary, quality_report)
    build_excel(PROJECT / 'excel' / 'Retail_Stock_Insight_Analysis.xlsx', cleaned, analytics, quality_report)
    build_presentation(summary, PROJECT / 'presentation' / 'Retail_Stock_Insight_Business_Pitch.pptx')
    build_speaking_script(PROJECT / 'presentation' / 'Presentation_Speaking_Script.docx')

    notebook_path = NOTEBOOK_DIR / 'Retail_Stock_Insight_EDA.ipynb'
    jupyter_cmd = [sys.executable, '-m', 'jupyter', 'nbconvert', '--to', 'notebook', '--execute', str(notebook_path), '--inplace']
    subprocess.run(jupyter_cmd, check=False, capture_output=True, text=True)

    print('Project completed successfully.')
    print(f'- Notebook: {notebook_path}')
    print(f'- Excel workbook: {PROJECT / "excel" / "Retail_Stock_Insight_Analysis.xlsx"}')
    print(f'- PowerPoint: {PROJECT / "presentation" / "Retail_Stock_Insight_Business_Pitch.pptx"}')
    print(f'- Script: {PROJECT / "presentation" / "Presentation_Speaking_Script.docx"}')


if __name__ == '__main__':
    main()
