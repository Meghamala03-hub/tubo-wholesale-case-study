import io
import re
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st

APP_TITLE = "Slow-Stock Action Tool"
AGING_COLS = [
    "Qty 0-30 days",
    "Qty 31-60 days",
    "Qty 61-90 days",
    "Qty 91-180 days",
    "Qty 180+ days",
]
AGED_COLS = ["Qty 91-180 days", "Qty 180+ days"]
REQUIRED = {
    "Item_Master": ["SKU", "Item Name", "Category", "Master Pack Qty", "Unit Cost (Rs)", "Alternate Item (name)", "Alternate Flag", "Status"],
    "Stock_Aging": ["SKU"] + AGING_COLS,
    "Sales_Lines": ["Month", "CustomerID", "CustomerName", "Industry", "SKU", "Qty", "Value (Rs)"],
    "Customers": ["CustomerID", "CustomerName", "Industry", "Area"],
}


def norm_text(value):
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def norm_sku(value):
    return norm_text(value).upper()


def norm_col(value):
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def normalize_columns(df: pd.DataFrame, table: str) -> pd.DataFrame:
    aliases = {norm_col(c): c for c in REQUIRED[table]}
    # Common variations in exports.
    extra = {
        "itemmaster": "Item_Master", "itemmasterdata": "Item_Master",
        "stockaging": "Stock_Aging", "saleslines": "Sales_Lines", "customer": "Customers",
    }
    rename = {}
    for col in df.columns:
        key = norm_col(col)
        if key in aliases:
            rename[col] = aliases[key]
        elif key in {"valueinrs", "valuerupees", "valuer s", "value"} and table == "Sales_Lines":
            rename[col] = "Value (Rs)"
    out = df.rename(columns=rename).copy()
    missing = [c for c in REQUIRED[table] if c not in out.columns]
    if missing:
        raise ValueError(f"{table} is missing required columns: {', '.join(missing)}")
    return out[REQUIRED[table]].copy()


def identify_table_from_filename(name: str):
    n = norm_col(Path(name).stem)
    if "itemmaster" in n or "1itemmaster" in n:
        return "Item_Master"
    if "stockaging" in n or "2stockaging" in n:
        return "Stock_Aging"
    if "saleslines" in n or "3saleslines" in n:
        return "Sales_Lines"
    if "customers" in n or "4customers" in n:
        return "Customers"
    return None


def load_uploaded_data(uploaded_files) -> Dict[str, pd.DataFrame]:
    tables = {}
    errors = []
    for file in uploaded_files:
        try:
            book = pd.ExcelFile(file)
            sheet_names = {norm_col(s): s for s in book.sheet_names}
            recognized = False
            for table in REQUIRED:
                key = norm_col(table)
                if key in sheet_names:
                    tables[table] = pd.read_excel(book, sheet_name=sheet_names[key])
                    recognized = True
            if recognized:
                continue
            table = identify_table_from_filename(file.name)
            if table:
                tables[table] = pd.read_excel(book, sheet_name=book.sheet_names[0])
            elif len(uploaded_files) == 1:
                raise ValueError("Could not identify the four required sheets. Upload the combined workbook or all four named files.")
        except Exception as exc:
            errors.append(f"{file.name}: {exc}")
    if errors:
        raise ValueError("\n".join(errors))
    missing_tables = [t for t in REQUIRED if t not in tables]
    if missing_tables:
        raise ValueError("Missing required data: " + ", ".join(missing_tables) + ". Upload the combined workbook or all four separate files.")
    return {t: normalize_columns(tables[t], t) for t in REQUIRED}


def clean_and_analyze(raw: Dict[str, pd.DataFrame], deduplicate: bool = True):
    issues = []
    items = raw["Item_Master"].copy()
    stock = raw["Stock_Aging"].copy()
    sales = raw["Sales_Lines"].copy()
    customers = raw["Customers"].copy()

    # Standardize string fields and SKU/customer keys.
    for col in ["SKU", "Item Name", "Category", "Alternate Item (name)", "Alternate Flag", "Status"]:
        items[col] = items[col].map(norm_text)
    for col in ["SKU"]:
        stock[col] = stock[col].map(norm_sku)
    for col in ["SKU", "CustomerID", "Month", "CustomerName", "Industry"]:
        sales[col] = sales[col].map(norm_text)
    sales["SKU"] = sales["SKU"].map(norm_sku)
    sales["CustomerID"] = sales["CustomerID"].map(norm_sku)
    for col in ["CustomerID", "CustomerName", "Industry", "Area"]:
        customers[col] = customers[col].map(norm_text)
    customers["CustomerID"] = customers["CustomerID"].map(norm_sku)
    items["SKU"] = items["SKU"].map(norm_sku)

    # Identify and normalize non-canonical SKU strings in sales.
    original_sales = raw["Sales_Lines"].copy()
    noncanonical_rows = int((original_sales["SKU"].astype(str) != original_sales["SKU"].astype(str).str.strip().str.upper()).sum())
    noncanonical_variants = int(original_sales.loc[original_sales["SKU"].astype(str) != original_sales["SKU"].astype(str).str.strip().str.upper(), "SKU"].nunique())
    if noncanonical_rows:
        issues.append({"Severity": "Medium", "Problem": f"{noncanonical_rows} sales rows have inconsistent SKU casing or whitespace ({noncanonical_variants} distinct raw SKU variants).", "Tool handling": "Trim whitespace and standardize SKU keys to uppercase before joins.", "Source fix requested": "Make the source system export canonical SKU codes and enforce a reference-data validation rule."})

    # Numeric conversion and missing-value checks.
    for col in ["Master Pack Qty", "Unit Cost (Rs)"]:
        items[col] = pd.to_numeric(items[col], errors="coerce")
    for col in AGING_COLS:
        stock[col] = pd.to_numeric(stock[col], errors="coerce")
    for col in ["Qty", "Value (Rs)"]:
        sales[col] = pd.to_numeric(sales[col], errors="coerce")

    # Exact duplicates are removed only when selected; always report them.
    duplicate_mask = sales.duplicated(subset=REQUIRED["Sales_Lines"], keep="first")
    duplicate_count = int(duplicate_mask.sum())
    if duplicate_count:
        issues.append({"Severity": "High", "Problem": f"{duplicate_count} exact duplicate sales rows were found.", "Tool handling": "When the recommended deduplication option is enabled, retain one copy of each exact row; the option can be switched off if the business confirms repeated rows are legitimate.", "Source fix requested": "Add a unique transaction/line ID and investigate why identical exports recur. Confirm these are duplicates before permanently deleting them."})
        if deduplicate:
            sales = sales.loc[~duplicate_mask].copy()

    # Master key checks.
    for table_name, df, key in [("Item Master", items, "SKU"), ("Stock Aging", stock, "SKU"), ("Customers", customers, "CustomerID")]:
        dup = int(df[key].duplicated(keep=False).sum())
        if dup:
            issues.append({"Severity": "High", "Problem": f"{table_name} contains duplicate {key} keys ({dup} rows involved).", "Tool handling": "Flag duplicates; aggregate stock/sales only where it is safe and avoid silently selecting a master record.", "Source fix requested": f"Enforce uniqueness for {key} in {table_name} and resolve conflicting records at source."})

    item_skus = set(items["SKU"])
    customer_ids = set(customers["CustomerID"])
    unknown_sales_skus = sorted(set(sales["SKU"]) - item_skus)
    if unknown_sales_skus:
        issues.append({"Severity": "High", "Problem": "Sales contain SKU codes not found in Item Master: " + ", ".join(unknown_sales_skus[:20]), "Tool handling": "Exclude unmatched SKU rows from item-level joins but retain them in the data-quality report; do not silently map them.", "Source fix requested": "Correct invalid SKU references or add the missing items to Item Master."})
    unknown_customer_ids = sorted(set(sales["CustomerID"]) - customer_ids)
    if unknown_customer_ids:
        issues.append({"Severity": "High", "Problem": "Sales contain customer IDs not found in Customers: " + ", ".join(unknown_customer_ids[:20]), "Tool handling": "Do not generate customer targets for unresolved IDs; flag the rows.", "Source fix requested": "Correct customer IDs or maintain a complete customer master."})

    missing_industry = customers[customers["Industry"].map(norm_text).eq("")]
    if len(missing_industry):
        ids = ", ".join(missing_industry["CustomerID"].tolist())
        issues.append({"Severity": "Medium", "Problem": f"Customer industry is missing for {ids}.", "Tool handling": "Customers without an industry are not eligible for Bucket 2 (Similar business), as specified in the brief.", "Source fix requested": "Populate the customer industry in the customer master and define a controlled industry list."})

    # Negative stock: preserve the source value but flag it; calculations use the reported signed quantities.
    negative_stock = stock[stock[AGING_COLS].lt(0).any(axis=1)]
    for _, row in negative_stock.iterrows():
        negs = [f"{c}={int(row[c]) if pd.notna(row[c]) and float(row[c]).is_integer() else row[c]}" for c in AGING_COLS if pd.notna(row[c]) and row[c] < 0]
        issues.append({"Severity": "High", "Problem": f"Negative stock quantity for {row['SKU']}: " + ", ".join(negs) + ".", "Tool handling": "Preserve the reported signed quantity for the calculation and flag the record; do not silently clamp it to zero.", "Source fix requested": "Reconcile warehouse count, stock movements, and aging-bucket allocation; correct the stock ledger at source."})

    # Active stock with no sales is a commercial exception, even when it is correctly classified as Red List.
    active_skus_for_check = set(items.loc[items["Status"].str.casefold().eq("active"), "SKU"])
    stock_totals_for_check = stock.assign(_total=stock[AGING_COLS].sum(axis=1)).groupby("SKU")["_total"].sum()
    sales_qty_by_sku_for_check = sales.groupby("SKU")["Qty"].sum()
    for sku in sorted(active_skus_for_check):
        stock_qty = float(stock_totals_for_check.get(sku, 0))
        sales_qty = float(sales_qty_by_sku_for_check.get(sku, 0))
        if stock_qty > 0 and sales_qty == 0:
            issues.append({"Severity": "Medium", "Problem": f"Active item {sku} has {stock_qty:g} units in stock but no sales in the six-month sales file.", "Tool handling": "Classify using the stated rules (normally Red List when sales are below one master pack) and keep the stock visible in the item list.", "Source fix requested": "Confirm the six-month sales extract is complete and check whether the stock should be promoted, returned, written off, or discontinued."})

    # Inactive item with stock or sales.
    inactive_skus = set(items.loc[items["Status"].str.casefold().eq("inactive"), "SKU"])
    stock_totals = stock.assign(_total=stock[AGING_COLS].sum(axis=1)).groupby("SKU")["_total"].sum()
    sales_skus = set(sales["SKU"])
    for sku in sorted(inactive_skus):
        qty = float(stock_totals.get(sku, 0))
        sold_rows = sales[sales["SKU"].eq(sku)]
        if qty != 0 or not sold_rows.empty:
            details = []
            if qty != 0:
                details.append(f"stock on hand {qty:g} units")
            if not sold_rows.empty:
                details.append(f"{len(sold_rows)} sales row(s), quantity {sold_rows['Qty'].sum():g}, value ₹{sold_rows['Value (Rs)'].sum():,.0f}")
            issues.append({"Severity": "High", "Problem": f"Inactive item {sku} still has " + " and ".join(details) + ".", "Tool handling": "Exclude the item from active classification and customer targeting; show the exception separately.", "Source fix requested": "Confirm whether the item should be reactivated, returned, written off, or cleared, and prevent sales against inactive SKUs."})

    # Alternate mapping based on item name. Only SWITCH is a merge instruction.
    name_to_sku = {norm_text(row["Item Name"]).casefold(): row["SKU"] for _, row in items.iterrows()}
    switch_map = {}
    for _, row in items.iterrows():
        flag = norm_text(row["Alternate Flag"]).upper()
        alt_name = norm_text(row["Alternate Item (name)"]).casefold()
        if alt_name and not flag:
            resolved = name_to_sku.get(alt_name)
            issues.append({"Severity": "Medium", "Problem": f"{row['SKU']} has an alternate-item name but no Alternate Flag" + (f"; it resolves to {resolved}." if resolved else " and the name does not resolve to a master item."), "Tool handling": "Do not merge because the business rule requires Alternate Flag = SWITCH; surface the metadata issue.", "Source fix requested": "Confirm whether the relationship is a pack-size switch and populate a valid flag, or clear an obsolete alternate reference."})
        if flag == "SWITCH":
            target_sku = name_to_sku.get(alt_name)
            if not target_sku:
                issues.append({"Severity": "High", "Problem": f"{row['SKU']} is flagged SWITCH but its alternate item name does not resolve to an Item Master SKU.", "Tool handling": "Do not merge an unresolved pair; flag the item and exclude it from normal switch handling until corrected.", "Source fix requested": "Replace free-text alternate names with a validated alternate SKU reference."})
            elif target_sku == row["SKU"]:
                issues.append({"Severity": "High", "Problem": f"{row['SKU']} points to itself as a SWITCH alternate.", "Tool handling": "Do not self-merge; flag for correction.", "Source fix requested": "Correct the alternate SKU reference."})
            else:
                switch_map[row["SKU"]] = target_sku
        elif flag and flag not in {"SWITCH"}:
            issues.append({"Severity": "Medium", "Problem": f"{row['SKU']} has unrecognized Alternate Flag '{flag}'.", "Tool handling": "Do not merge because only SWITCH is defined as a merge instruction in this case.", "Source fix requested": "Document the allowed flag values and use a controlled dropdown/reference list."})

    # Data-quality flags for non-positive pack/cost or missing numeric values.
    for _, row in items[items["Master Pack Qty"].isna() | items["Master Pack Qty"].le(0)].iterrows():
        issues.append({"Severity": "High", "Problem": f"{row['SKU']} has missing or non-positive Master Pack Qty.", "Tool handling": "Do not classify reliably until pack quantity is corrected; show the issue.", "Source fix requested": "Maintain a positive numeric master pack quantity for every active SKU."})
    for _, row in items[items["Unit Cost (Rs)"].isna() | items["Unit Cost (Rs)"].lt(0)].iterrows():
        issues.append({"Severity": "High", "Problem": f"{row['SKU']} has missing or negative Unit Cost.", "Tool handling": "Aged value cannot be trusted for this SKU; flag it for correction.", "Source fix requested": "Maintain a valid unit cost and define the valuation date/source."})
    for _, row in stock[stock[AGING_COLS].isna().any(axis=1)].iterrows():
        issues.append({"Severity": "High", "Problem": f"Stock aging has missing bucket quantities for {row['SKU']}.", "Tool handling": "Flag missing quantities rather than treating them as verified zeroes.", "Source fix requested": "Ensure every SKU has a numeric value in each aging bucket."})
    for _, row in sales[sales[["Qty", "Value (Rs)"]].isna().any(axis=1)].iterrows():
        issues.append({"Severity": "High", "Problem": f"Sales row for {row['SKU']} / {row['CustomerID']} has missing or non-numeric quantity/value.", "Tool handling": "Exclude invalid rows from numeric totals and report them.", "Source fix requested": "Require valid numeric quantity and sales value in every transaction line."})

    # Clean stock and aggregate across switch pairs.
    stock = stock.groupby("SKU", as_index=False)[AGING_COLS].sum(min_count=1)
    stock["Source SKU"] = stock["SKU"]
    stock["SKU"] = stock["SKU"].replace(switch_map)
    stock_agg = stock.groupby("SKU", as_index=False)[AGING_COLS].sum(min_count=1)

    # Map sales switch partner to the new SKU and retain original SKU for bought-before checks.
    sales["Original SKU"] = sales["SKU"]
    sales["Mapped SKU"] = sales["SKU"].replace(switch_map)
    sales = sales.dropna(subset=["Qty", "Value (Rs)"]).copy()

    sales_agg = sales.groupby("Mapped SKU", as_index=False).agg(
        SalesQty=("Qty", "sum"), SalesValue=("Value (Rs)", "sum"), SalesLines=("Mapped SKU", "size")
    ).rename(columns={"Mapped SKU": "SKU"})

    # Only active, non-source SWITCH items appear in the classification list.
    source_switch_skus = set(switch_map.keys())
    active = items[items["Status"].str.casefold().eq("active") & ~items["SKU"].isin(source_switch_skus)].copy()
    # Keep rows with unresolved SWITCH mappings visible; their status is not silently merged.
    active = active.merge(stock_agg, on="SKU", how="left", validate="one_to_one")
    active = active.merge(sales_agg, on="SKU", how="left", validate="one_to_one")
    for c in AGING_COLS:
        active[c] = pd.to_numeric(active[c], errors="coerce")
    active["SalesQty"] = active["SalesQty"].fillna(0)
    active["SalesValue"] = active["SalesValue"].fillna(0)
    active["SalesLines"] = active["SalesLines"].fillna(0).astype(int)
    active["HasMissingStockBucket"] = active[AGING_COLS].isna().any(axis=1)
    active["TotalStock"] = active[AGING_COLS].sum(axis=1, min_count=1)
    active["AgedQty"] = active[AGED_COLS].sum(axis=1, min_count=1)
    active["AvgMonthlySales"] = active["SalesQty"] / 6
    active["MonthsToClear"] = active["TotalStock"] / active["AvgMonthlySales"].replace(0, pd.NA)
    active["AgedValue"] = active["AgedQty"] * active["Unit Cost (Rs)"]

    def classify(row):
        if pd.isna(row["Master Pack Qty"]) or row["Master Pack Qty"] <= 0 or pd.isna(row["TotalStock"]) or pd.isna(row["AgedQty"]) or row["HasMissingStockBucket"]:
            return "Review Data"
        if row["SalesQty"] < row["Master Pack Qty"]:
            return "Red List"
        if row["AgedQty"] <= 0 or (pd.notna(row["MonthsToClear"]) and row["MonthsToClear"] < 3):
            return "No Action"
        return "Push"

    active["Class"] = active.apply(classify, axis=1)
    # Preserve infinity for zero sales in the presentation, and provide readable columns.
    active["MonthsToClear"] = active["MonthsToClear"].astype("Float64")

    # Customer total spend is calculated over the six-month sales data after the selected deduplication.
    cust_totals = sales.groupby("CustomerID", as_index=False).agg(CustomerPurchaseValue=("Value (Rs)", "sum"))
    customer_view = customers.merge(cust_totals, on="CustomerID", how="left")
    customer_view["CustomerPurchaseValue"] = customer_view["CustomerPurchaseValue"].fillna(0)
    customer_view["Industry"] = customer_view["Industry"].map(norm_text)

    # Categories and buyer industries use the master data and the mapped SKU.
    category_by_sku = items[items["Status"].str.casefold().eq("active") & ~items["SKU"].isin(source_switch_skus)].set_index("SKU")["Category"].to_dict()
    customer_industry = customers.set_index("CustomerID")["Industry"].to_dict()
    customer_area = customers.set_index("CustomerID")["Area"].to_dict()
    customer_name = customers.set_index("CustomerID")["CustomerName"].to_dict()
    targets = []
    for _, item in active[active["Class"].eq("Push")].iterrows():
        sku = item["SKU"]
        item_sales = sales[sales["Mapped SKU"].eq(sku)]
        buyers = set(item_sales["CustomerID"])
        buyer_industries = {norm_text(customer_industry.get(cid, "")).casefold() for cid in buyers if norm_text(customer_industry.get(cid, ""))}
        category = item["Category"]
        category_skus = {s for s, cat in category_by_sku.items() if norm_text(cat).casefold() == norm_text(category).casefold() and s != sku}
        category_buyers = set(sales.loc[sales["Mapped SKU"].isin(category_skus), "CustomerID"])
        for _, cust in customer_view.iterrows():
            cid = cust["CustomerID"]
            if cust["CustomerPurchaseValue"] <= 100000:
                continue
            if cid in buyers:
                bucket = 1
                bucket_label = "1 – Bought before"
            elif norm_text(cust["Industry"]) and norm_text(cust["Industry"]).casefold() in buyer_industries:
                bucket = 2
                bucket_label = "2 – Similar business"
            elif cid in category_buyers:
                bucket = 3
                bucket_label = "3 – Same category"
            else:
                continue
            targets.append({
                "SKU": sku, "Item Name": item["Item Name"], "Category": category,
                "CustomerID": cid, "CustomerName": cust["CustomerName"], "Industry": cust["Industry"],
                "Area": cust["Area"], "CustomerPurchaseValue": cust["CustomerPurchaseValue"],
                "Bucket": bucket, "Bucket Label": bucket_label,
            })
    target_df = pd.DataFrame(targets, columns=["SKU", "Item Name", "Category", "CustomerID", "CustomerName", "Industry", "Area", "CustomerPurchaseValue", "Bucket", "Bucket Label"])
    if not target_df.empty:
        target_df = target_df.sort_values(["SKU", "Bucket", "CustomerPurchaseValue", "CustomerID"], ascending=[True, True, False, True]).reset_index(drop=True)

    # Surface data rows for inactive items, including their stock and sales.
    inactive_rows = []
    for _, row in items[items["Status"].str.casefold().eq("inactive")].iterrows():
        sku = row["SKU"]
        stock_row = stock[stock["SKU"].eq(sku)]
        total_stock = float(stock_row[AGING_COLS].sum(axis=1).sum()) if not stock_row.empty else 0
        sold = sales[sales["SKU"].eq(sku)]
        if total_stock or not sold.empty:
            inactive_rows.append({"SKU": sku, "Item Name": row["Item Name"], "Stock Qty": total_stock, "Sales Qty": sold["Qty"].sum(), "Sales Value (Rs)": sold["Value (Rs)"].sum(), "Status": row["Status"]})
    inactive_df = pd.DataFrame(inactive_rows)

    issue_df = pd.DataFrame(issues, columns=["Severity", "Problem", "Tool handling", "Source fix requested"])
    return active, target_df, issue_df, inactive_df, sales, switch_map, customer_view


def make_export(summary_df, items_df, targets_df, issues_df, inactive_df):
    buff = io.BytesIO()
    with pd.ExcelWriter(buff, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        items_df.to_excel(writer, sheet_name="Item Analysis", index=False)
        targets_df.to_excel(writer, sheet_name="Customer Targets", index=False)
        issues_df.to_excel(writer, sheet_name="Data Quality", index=False)
        inactive_df.to_excel(writer, sheet_name="Inactive Exceptions", index=False)
    buff.seek(0)
    return buff


st.set_page_config(page_title=APP_TITLE, page_icon="📦", layout="wide")
st.title("📦 Slow-Stock Action Tool")
st.caption("Monthly slow-stock review · Company X case study · Sample data is synthetic")
with st.expander("How the tool classifies items", expanded=False):
    st.markdown("""
    1. **Pack-size switches:** combine the source and alternate SKU only when `Alternate Flag = SWITCH`; report under the alternate/new SKU.
    2. **Inactive items:** exclude from active classification and report separately if stock or sales remain.
    3. **Aged stock:** quantity in the 91–180 and 180+ day buckets; aged value = aged quantity × unit cost.
    4. **Months to clear:** total stock ÷ (six-month sales quantity ÷ 6).
    5. **Class order:** Red List if six-month sales are less than one master pack; otherwise No Action if no aged stock or months to clear < 3; otherwise Push.
    6. **Customer targets:** six-month customer spend must be **greater than ₹1,00,000**. Bucket priority is 1, then 2, then 3; within a bucket, higher spend ranks first.
    """)

uploaded = st.file_uploader("Upload the combined workbook or all four separate Excel files", type=["xlsx", "xls"], accept_multiple_files=True)
deduplicate = st.checkbox("Remove exact duplicate sales rows (recommended; review data-quality warning)", value=True)
if not uploaded:
    st.info("Upload the provided `Case_Data_Technology_and_Data_Management_All_Sheets.xlsx` workbook to begin. You can also upload the four named files together.")
    st.stop()

try:
    raw_data = load_uploaded_data(uploaded)
    item_df, target_df, issues_df, inactive_df, cleaned_sales, switch_map, customer_view = clean_and_analyze(raw_data, deduplicate=deduplicate)
except Exception as e:
    st.error(f"Could not process the files: {e}")
    st.stop()

# Summary cards
counts = item_df["Class"].value_counts().to_dict()
push_value = item_df.loc[item_df["Class"].eq("Push"), "AgedValue"].sum(skipna=True)
c1, c2, c3, c4 = st.columns(4)
c1.metric("Active items analyzed", f"{len(item_df):,}")
c2.metric("Push", f"{counts.get('Push', 0):,}")
c3.metric("No Action", f"{counts.get('No Action', 0):,}")
c4.metric("Red List", f"{counts.get('Red List', 0):,}")
st.metric("Total aged value in Push items", f"₹{push_value:,.0f}")

if not issues_df.empty:
    st.warning(f"Data-quality checks found {len(issues_df)} issue(s). Review the Data Quality tab before acting on recommendations.")

summary = pd.DataFrame([
    {"Metric": "Active items analyzed", "Value": len(item_df)},
    {"Metric": "Red List items", "Value": counts.get("Red List", 0)},
    {"Metric": "No Action items", "Value": counts.get("No Action", 0)},
    {"Metric": "Push items", "Value": counts.get("Push", 0)},
    {"Metric": "Review Data items", "Value": counts.get("Review Data", 0)},
    {"Metric": "Total aged value in Push items (Rs)", "Value": push_value},
])

tab1, tab2, tab3, tab4, tab5 = st.tabs(["Overview", "Item List", "Customer Targets", "Data Quality", "Inactive Exceptions"])
with tab1:
    st.subheader("Class summary")
    st.dataframe(summary, use_container_width=True, hide_index=True)
    st.subheader("Top 10 Push items by aged value")
    top10 = item_df[item_df["Class"].eq("Push")].sort_values("AgedValue", ascending=False).head(10)
    show_cols = ["SKU", "Item Name", "Category", "TotalStock", "AgedQty", "Unit Cost (Rs)", "AgedValue", "SalesQty", "MonthsToClear"]
    st.dataframe(top10[show_cols].rename(columns={"TotalStock": "Total Stock", "AgedQty": "Aged Qty", "SalesQty": "6M Sales Qty", "MonthsToClear": "Months to Clear"}), use_container_width=True, hide_index=True)
    st.caption("Months to clear is shown as ∞/missing when six-month sales are zero. Red List takes precedence over No Action and Push.")
with tab2:
    st.subheader("Filterable item analysis")
    f1, f2 = st.columns(2)
    classes = ["All"] + sorted(item_df["Class"].dropna().unique().tolist())
    categories = ["All"] + sorted(item_df["Category"].dropna().unique().tolist())
    selected_class = f1.selectbox("Class", classes)
    selected_category = f2.selectbox("Category", categories)
    filtered = item_df.copy()
    if selected_class != "All":
        filtered = filtered[filtered["Class"] == selected_class]
    if selected_category != "All":
        filtered = filtered[filtered["Category"] == selected_category]
    cols = ["SKU", "Item Name", "Category", "Class", "Master Pack Qty", "Unit Cost (Rs)", "SalesQty", "AvgMonthlySales", "TotalStock", "AgedQty", "AgedValue", "MonthsToClear"]
    st.dataframe(filtered[cols].rename(columns={"SalesQty": "6M Sales Qty", "AvgMonthlySales": "Avg Monthly Sales", "TotalStock": "Total Stock", "AgedQty": "Aged Qty", "AgedValue": "Aged Value (Rs)", "MonthsToClear": "Months to Clear"}), use_container_width=True, hide_index=True)
with tab3:
    st.subheader("Customer call list")
    push_items = item_df[item_df["Class"].eq("Push")].sort_values("SKU")
    if push_items.empty:
        st.info("No Push items found in this run.")
    else:
        selected_sku = st.selectbox("Choose a Push item", push_items["SKU"].tolist(), format_func=lambda s: f"{s} — {push_items.loc[push_items['SKU'].eq(s), 'Item Name'].iloc[0]}")
        chosen = target_df[target_df["SKU"].eq(selected_sku)].sort_values(["Bucket", "CustomerPurchaseValue"], ascending=[True, False]) if not target_df.empty else target_df
        if chosen.empty:
            st.info("No eligible customer targets met the business rules for this item.")
        else:
            st.caption("Call customers in the displayed order. This view is designed to stay simple on mobile screens.")
            for i, row in enumerate(chosen.to_dict("records"), start=1):
                with st.container(border=True):
                    st.markdown(f"**{i}. {row['CustomerName']}** · {row['CustomerID']}")
                    st.write(f"Bucket: {row['Bucket Label']}  |  Industry: {row['Industry'] or 'Not recorded'}  |  Area: {row['Area']}")
                    st.write(f"Six-month purchase value: ₹{row['CustomerPurchaseValue']:,.0f}")
        st.subheader("Downloadable target rows")
        st.dataframe(chosen, use_container_width=True, hide_index=True)
with tab4:
    st.subheader("Data-quality findings and handling")
    if issues_df.empty:
        st.success("No data-quality issues detected by the current checks.")
    else:
        st.dataframe(issues_df, use_container_width=True, hide_index=True)
    st.markdown("**Important assumption:** exact duplicate sales rows are removed only when the checkbox is enabled. Confirm them against transaction IDs before treating deduplication as a permanent source-system correction.")
with tab5:
    st.subheader("Inactive items with stock or sales")
    if inactive_df.empty:
        st.success("No inactive items with stock or sales were found.")
    else:
        st.dataframe(inactive_df, use_container_width=True, hide_index=True)

export = make_export(summary, item_df, target_df, issues_df, inactive_df)
st.download_button("⬇️ Export results to Excel", data=export, file_name="slow_stock_review_results.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
st.caption("Before using the list commercially, review all data-quality exceptions and confirm source-data assumptions.")
