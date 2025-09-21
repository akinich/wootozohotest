import streamlit as st
import pandas as pd
import requests
from datetime import datetime
from dateutil.parser import parse  # safer date parsing
from collections import Counter
from io import BytesIO
from zipfile import ZipFile
import os

# ------------------------
# WooCommerce API settings
WC_API_URL = st.secrets.get("WC_API_URL")
WC_CONSUMER_KEY = st.secrets.get("WC_CONSUMER_KEY")
WC_CONSUMER_SECRET = st.secrets.get("WC_CONSUMER_SECRET")

if not WC_API_URL or not WC_CONSUMER_KEY or not WC_CONSUMER_SECRET:
    st.error("WooCommerce API credentials are missing. Please add them to Streamlit secrets.")
    st.stop()

# ------------------------
st.title("WooCommerce → Accounting CSV/Excel Export Tool")

# Date input fields
start_date = st.date_input("Start Date")
end_date = st.date_input("End Date")

# Invoice number customization
invoice_prefix = st.text_input("Invoice Prefix", value="ECHE/2526/")
start_sequence = st.number_input("Starting Sequence Number", min_value=1, value=608)

if start_date > end_date:
    st.error("Start date cannot be after end date.")

fetch_button = st.button("Fetch Orders", disabled=(start_date > end_date))

def to_float(x):
    try:
        if x is None or x == "":
            return 0.0
        return float(x)
    except Exception:
        return 0.0

# ------------------------
# Load item database for CSV mapping
ITEM_DB_PATH = "Item database.xlsx"
item_lookup = {}
if os.path.exists(ITEM_DB_PATH):
    item_db_df = pd.read_excel(ITEM_DB_PATH, dtype=str)  # HSN as string
    for _, row in item_db_df.iterrows():
        woocommerce_name = str(row.get("woocommerce name", "")).strip()
        item_lookup[woocommerce_name] = {
            "zoho_name": str(row.get("zoho name", "")).strip(),
            "hsn": str(row.get("hsn", "")).strip(),
            "usage_unit": str(row.get("usage count", "")).strip()
        }
else:
    st.warning(f"Item database not found at {ITEM_DB_PATH}. Item mapping will be skipped.")

# ------------------------
if fetch_button:
    st.info("Preparing to fetch orders from WooCommerce...")

    # Convert to WooCommerce ISO format
    start_iso = start_date.strftime("%Y-%m-%dT00:00:00")
    end_iso = end_date.strftime("%Y-%m-%dT23:59:59")

    # Pagination loop - Fetch ALL orders in range
    all_orders = []
    page = 1
    try:
        with st.spinner("Fetching orders from WooCommerce..."):
            while True:
                response = requests.get(
                    f"{WC_API_URL}/orders",
                    params={
                        "after": start_iso,
                        "before": end_iso,
                        "per_page": 100,
                        "page": page
                    },
                    auth=(WC_CONSUMER_KEY, WC_CONSUMER_SECRET),
                    timeout=30
                )
                response.raise_for_status()
                orders = response.json()
                if not orders:
                    break
                all_orders.extend(orders)
                page += 1
    except requests.exceptions.RequestException as e:
        st.error(f"Error fetching orders from WooCommerce: {e}")
        st.stop()

    if not all_orders:
        st.warning("No orders found in this date range.")
        st.stop()

    # Sort ascending by order ID for invoice sequence
    all_orders.sort(key=lambda x: x["id"])

    # ------------------------
    # Count orders by status
    status_counts = Counter(order["status"].lower() for order in all_orders)
    def get_status_count(variants):
        return sum(status_counts.get(v, 0) for v in variants)

    # ------------------------
    # Transform only COMPLETED orders into CSV rows
    csv_rows = []
    sequence_number = start_sequence

    for order in all_orders:
        if order["status"].lower() != "completed":
            continue  # only export completed orders to CSV

        order_id = order["id"]
        invoice_number = f"{invoice_prefix}{sequence_number:05d}"
        sequence_number += 1

        # Safe date parsing
        invoice_date = parse(order["date_created"]).strftime("%Y-%m-%d %H:%M:%S")

        customer_name = f"{order['billing'].get('first_name','')} {order['billing'].get('last_name','')}".strip()
        place_of_supply = order['billing'].get('state', '')
        currency = order.get('currency', '')
        shipping_charge = to_float(order.get('shipping_total', 0))
        entity_discount = to_float(order.get('discount_total', 0))

        for item in order.get("line_items", []):
            # Pull product metadata
            product_meta = item.get("meta_data", []) or []
            hsn = ""
            usage_unit = ""
            for meta in product_meta:
                key = str(meta.get("key", "")).lower()
                if key == "hsn":
                    hsn = meta.get("value", "")
                if key == "usage unit":
                    usage_unit = meta.get("value", "")

            item_name = item.get("name", "")

            # ------------------------
            # Map from Item database if available
            if item_name in item_lookup:
                item_name = item_lookup[item_name]["zoho_name"] or item_name
                hsn = item_lookup[item_name]["hsn"] or hsn
                usage_unit = item_lookup[item_name]["usage_unit"] or usage_unit

            row = {
                "Invoice Number": invoice_number,
                "PurchaseOrder": order_id,
                "Invoice Date": invoice_date,
                "Invoice Status": order["status"].capitalize(),
                "Customer Name": customer_name,
                "Place of Supply": place_of_supply,
                "Currency Code": currency,
                "Item Name": item_name,
                "HSN/SAC": hsn,
                "Item Type": item.get("type", "goods"),
                "Quantity": item.get("quantity", 0),
                "Usage unit": usage_unit,
                "Item Price": item.get("price", ""),
                "Is Inclusive Tax": "FALSE",
                "Item Tax %": item.get("tax_class") or "0",
                "Discount Type": "entity_level",
                "Is Discount Before Tax": "TRUE",
                "Entity Discount Amount": entity_discount,
                "Shipping Charge": shipping_charge,
                "Item Tax Exemption Reason": "ITEM EXEMPT FROM GST",
                "Supply Type": "Exempted",
                "GST Treatment": "consumer"
            }
            csv_rows.append(row)

    df = pd.DataFrame(csv_rows)
    st.dataframe(df.head(50))

    # ------------------------
    # Revenue calculation only from WooCommerce order totals
    completed_orders = [o for o in all_orders if o["status"].lower() == "completed"]
    total_revenue_by_order_total = 0.0
    for order in completed_orders:
        order_total = to_float(order.get("total", 0))
        refunds = order.get("refunds") or []
        refund_total = sum(to_float(r.get("amount") or r.get("total") or r.get("refund_total") or 0) for r in refunds)
        total_revenue_by_order_total += order_total - refund_total

    # ------------------------
    # Summary metrics
    first_order_id = completed_orders[0]["id"] if completed_orders else None
    last_order_id = completed_orders[-1]["id"] if completed_orders else None
    first_invoice_number = f"{invoice_prefix}{start_sequence:05d}"
    last_invoice_number = f"{invoice_prefix}{sequence_number - 1:05d}" if completed_orders else None

    with st.expander("View Summary Report"):
        st.subheader("Summary Report")
        st.write(f"**Total Orders Fetched:** {len(all_orders)}")
        st.write("---")
        st.write("### Orders by Status")
        st.write(f"- Completed: **{get_status_count(['completed'])}**")
        st.write(f"- Processing: **{get_status_count(['processing'])}**")
        st.write(f"- On Hold: **{get_status_count(['on-hold','on_hold','on hold'])}**")
        st.write(f"- Cancelled: **{get_status_count(['cancelled','canceled'])}**")
        st.write(f"- Pending Payment: **{get_status_count(['pending','pending payment','pending-payment'])}**")
        st.write("---")
        if completed_orders:
            st.write(f"**Completed Order ID Range:** {first_order_id} → {last_order_id}")
            st.write(f"**Invoice Number Range:** {first_invoice_number} → {last_invoice_number}")
        st.write("---")
        st.write(f"**Total Revenue (WooCommerce order totals, net of refunds):** ₹ {total_revenue_by_order_total:,.2f}")

    # ------------------------
    # Generate summary CSV for download
    summary_csv_data = []
    for order in completed_orders:
        order_id = order["id"]
        invoice_number = f"{invoice_prefix}{start_sequence:05d}"
        start_sequence += 1
        invoice_date = parse(order["date_created"]).strftime("%Y-%m-%d %H:%M:%S")
        customer_name = f"{order['billing'].get('first_name','')} {order['billing'].get('last_name','')}".strip()
        order_total = to_float(order.get("total", 0))
        summary_csv_data.append({
            "Invoice Number": invoice_number,
            "Order Number": order_id,
            "Invoice Date": invoice_date,
            "Customer Name": customer_name,
            "Order Total": order_total
        })
    summary_df = pd.DataFrame(summary_csv_data)

    # ------------------------
    # Generate Excel file (optional)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        summary_df.to_excel(writer, index=False, sheet_name="Summary Metrics")
        df.to_excel(writer, index=False, sheet_name="Order Details")
    excel_data = output.getvalue()

    # ------------------------
    # ZIP both files for single download
    zip_buffer = BytesIO()
    with ZipFile(zip_buffer, "w") as zip_file:
        zip_file.writestr(f"orders_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.csv",
                          df.to_csv(index=False).encode('utf-8'))
        zip_file.writestr(f"summary_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx",
                          excel_data)
    zip_buffer.seek(0)

    st.download_button(
        label="Download CSV + Excel (ZIP)",
        data=zip_buffer,
        file_name=f"woocommerce_export_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.zip",
        mime="application/zip"
    )