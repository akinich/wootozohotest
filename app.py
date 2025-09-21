import streamlit as st
import pandas as pd
import requests
from datetime import datetime
from dateutil.parser import parse  # safer date parsing

# ------------------------
# WooCommerce API settings (use Streamlit secrets for production)
WC_API_URL = st.secrets.get("WC_API_URL")
WC_CONSUMER_KEY = st.secrets.get("WC_CONSUMER_KEY")
WC_CONSUMER_SECRET = st.secrets.get("WC_CONSUMER_SECRET")

# Validate secrets to avoid using hardcoded defaults
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

# Validate date range
if start_date > end_date:
    st.error("Start date cannot be after end date.")

fetch_button = st.button("Fetch Orders", disabled=(start_date > end_date))

if fetch_button:
    st.info("Preparing to fetch completed orders from WooCommerce...")

    # Convert to WooCommerce ISO format
    start_iso = start_date.strftime("%Y-%m-%dT00:00:00")
    end_iso = end_date.strftime("%Y-%m-%dT23:59:59")

    # Pagination loop
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
                        "status": "completed",
                        "per_page": 100,
                        "page": page
                    },
                    auth=(WC_CONSUMER_KEY, WC_CONSUMER_SECRET),
                    timeout=30  # avoid hanging
                )
                response.raise_for_status()  # raise if HTTP error occurs
                orders = response.json()

                if not orders:
                    break  # no more pages
                all_orders.extend(orders)
                page += 1

    except requests.exceptions.RequestException as e:
        st.error(f"Error fetching orders from WooCommerce: {e}")
        st.stop()

    # Stop if no orders found
    if not all_orders:
        st.warning("No completed orders found in this date range.")
        st.stop()

    # --- Sort ascending by order ID for invoice sequence ---
    all_orders.sort(key=lambda x: x["id"])

    # ------------------------
    # Transform WooCommerce orders into CSV rows
    csv_rows = []
    sequence_number = start_sequence

    for order in all_orders:
        order_id = order["id"]
        invoice_number = f"{invoice_prefix}{sequence_number:05d}"
        sequence_number += 1

        # Safe date parsing (handles timezone)
        invoice_date = parse(order["date_created"]).strftime("%Y-%m-%d %H:%M:%S")

        customer_name = f"{order['billing']['first_name']} {order['billing']['last_name']}".strip()
        place_of_supply = order['billing']['state']
        currency = order['currency']
        shipping_charge = float(order['shipping_total']) if order['shipping_total'] else 0
        entity_discount = float(order['discount_total']) if order['discount_total'] else 0

        for item in order["line_items"]:
            # Pull product metadata for HSN and usage unit
            product_meta = item.get("meta_data", [])
            hsn = ""
            usage_unit = ""
            for meta in product_meta:
                if meta["key"].lower() == "hsn":
                    hsn = meta["value"]
                if meta["key"].lower() == "usage unit":
                    usage_unit = meta["value"]

            row = {
                "Invoice Number": invoice_number,
                "PurchaseOrder": order_id,
                "Invoice Date": invoice_date,
                "Invoice Status": order["status"].capitalize(),
                "Customer Name": customer_name,
                "Place of Supply": place_of_supply,
                "Currency Code": currency,
                "Item Name": item["name"],
                "HSN/SAC": hsn,
                "Item Type": item.get("type", "goods"),
                "Quantity": item["quantity"],
                "Usage unit": usage_unit,
                "Item Price": item["price"],
                "Is Inclusive Tax": "FALSE",
                "Item Tax %": item.get("tax_class") or "0",  # safe default
                "Discount Type": "entity_level",
                "Is Discount Before Tax": "TRUE",
                "Entity Discount Amount": entity_discount,
                "Shipping Charge": shipping_charge,
                "Item Tax Exemption Reason": "ITEM EXEMPT FROM GST",
                "Supply Type": "Exempted",
                "GST Treatment": "consumer"
            }
            csv_rows.append(row)

    # Create DataFrame
    df = pd.DataFrame(csv_rows)

    # Limit preview for performance
    st.dataframe(df.head(50))

    # ------------------------
    # Summary report
    first_order_id = all_orders[0]["id"]
    last_order_id = all_orders[-1]["id"]
    first_invoice_number = f"{invoice_prefix}{start_sequence:05d}"
    last_invoice_number = f"{invoice_prefix}{sequence_number - 1:05d}"

    with st.expander("View Summary Report"):
        st.subheader("Summary Report")
        st.write(f"**Total Orders Processed:** {len(all_orders)}")
        st.write(f"**Order IDs:** {first_order_id} → {last_order_id}")
        st.write(f"**Invoice Numbers:** {first_invoice_number} → {last_invoice_number}")

    # ------------------------
    # CSV download
    csv_bytes = df.to_csv(index=False).encode('utf-8')
    download_filename = f"orders_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.csv"

    st.download_button(
        label="Download CSV",
        data=csv_bytes,
        file_name=download_filename,
        mime="text/csv"
    )
