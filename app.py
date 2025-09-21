import streamlit as st
import pandas as pd
import requests
from datetime import datetime
from dateutil.parser import parse  # safer date parsing
from collections import Counter
from io import BytesIO

# ------------------------
# WooCommerce API settings
WC_API_URL = st.secrets.get("WC_API_URL")
WC_CONSUMER_KEY = st.secrets.get("WC_CONSUMER_KEY")
WC_CONSUMER_SECRET = st.secrets.get("WC_CONSUMER_SECRET")

if not WC_API_URL or not WC_CONSUMER_KEY or not WC_CONSUMER_SECRET:
    st.error("WooCommerce API credentials are missing. Please add them to Streamlit secrets.")
    st.stop()

# ------------------------
st.title("WooCommerce → Accounting CSV Export Tool")

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

            item_price = item.get("price", "")
            row = {
                "Invoice Number": invoice_number,
                "PurchaseOrder": order_id,
                "Invoice Date": invoice_date,
                "Invoice Status": order["status"].capitalize(),
                "Customer Name": customer_name,
                "Place of Supply": place_of_supply,
                "Currency Code": currency,
                "Item Name": item.get("name", ""),
                "HSN/SAC": hsn,
                "Item Type": item.get("type", "goods"),
                "Quantity": item.get("quantity", 0),
                "Usage unit": usage_unit,
                "Item Price": item_price,
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

        # Handle refunds
        refunds = order.get("refunds") or []
        refund_total = sum(to_float(r.get("amount") or r.get("total") or r.get("refund_total") or 0) for r in refunds)

        net_order_total = order_total - refund_total
        total_revenue_by_order_total += net_order_total

    # ------------------------
    # Summary metrics
    first_order_id = completed_orders[0]["id"] if completed_orders else None
    last_order_id = completed_orders[-1]["id"] if completed_orders else None
    first_invoice_number = f"{invoice_prefix}{start_sequence:05d}"
    last_invoice_number = f"{invoice_prefix}{sequence_number - 1:05d}" if completed_orders else None

    summary_metrics = {
        "Metric": [
            "Total Orders Fetched",
            "Completed Orders",
            "Processing Orders",
            "On Hold Orders",
            "Cancelled Orders",
            "Pending Payment Orders",
            "Completed Order ID Range",
            "Invoice Number Range",
            "Total Revenue (Net of Refunds)"
        ],
        "Value": [
            len(all_orders),
            get_status_count(['completed']),
            get_status_count(['processing']),
            get_status_count(['on-hold','on_hold','on hold']),
            get_status_count(['cancelled','canceled']),
            get_status_count(['pending','pending payment','pending-payment']),
            f"{first_order_id} → {last_order_id}" if completed_orders else "",
            f"{first_invoice_number} → {last_invoice_number}" if completed_orders else "",
            f"₹ {total_revenue_by_order_total:,.2f}"
        ]
    }
    summary_df = pd.DataFrame(summary_metrics)

    # ------------------------
    # Per-order details sheet
    order_details_rows = []
    sequence_number_temp = start_sequence
    for order in completed_orders:
        invoice_number_temp = f"{invoice_prefix}{sequence_number_temp:05d}"
        sequence_number_temp += 1
        order_total = to_float(order.get("total", 0))
        refunds = order.get("refunds") or []
        refund_total = sum(to_float(r.get("amount") or r.get("total") or r.get("refund_total") or 0) for r in refunds)
        net_total = order_total - refund_total

        order_details_rows.append({
            "Invoice Number": invoice_number_temp,
            "Order Number": order["id"],
            "Date": parse(order["date_created"]).strftime("%Y-%m-%d %H:%M:%S"),
            "Customer Name": f"{order['billing'].get('first_name','')} {order['billing'].get('last_name','')}".strip(),
            "Order Total": net_total
        })

    order_details_df = pd.DataFrame(order_details_rows)

    # ------------------------
    # Export to Excel with two sheets
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        summary_df.to_excel(writer, index=False, sheet_name="Summary Metrics")
        order_details_df.to_excel(writer, index=False, sheet_name="Order Details")
    excel_data = output.getvalue()

    st.download_button(
        label="Download Summary Report (Excel)",
        data=excel_data,
        file_name=f"summary_report_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # ------------------------
    # CSV download for accounting
    csv_bytes = df.to_csv(index=False).encode('utf-8')
    download_filename = f"orders_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.csv"

    st.download_button(
        label="Download CSV (Completed Orders Only)",
        data=csv_bytes,
        file_name=download_filename,
        mime="text/csv"
    )
