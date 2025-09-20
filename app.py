# app.py
import streamlit as st
import pandas as pd
import requests
from datetime import datetime

# ------------------------
# Streamlit secrets (replace with your actual keys or use Streamlit secrets management)
WC_API_URL = st.secrets.get("WC_API_URL", "https://yourstore.com/wp-json/wc/v3")
WC_CONSUMER_KEY = st.secrets.get("WC_CONSUMER_KEY", "ck_xxxxx")
WC_CONSUMER_SECRET = st.secrets.get("WC_CONSUMER_SECRET", "cs_xxxxx")

# ------------------------
st.title("WooCommerce → Accounting CSV Export Tool")

# Date input
start_date = st.date_input("Start Date")
end_date = st.date_input("End Date")

if start_date > end_date:
    st.error("Start date cannot be after end date.")

fetch_button = st.button("Fetch Orders")

if fetch_button:
    st.info("Fetching orders from WooCommerce...")
    
    # Convert dates to ISO format for WooCommerce API
    start_iso = start_date.strftime("%Y-%m-%dT00:00:00")
    end_iso = end_date.strftime("%Y-%m-%dT23:59:59")

    # WooCommerce API endpoint for orders
    orders_endpoint = f"{WC_API_URL}/orders"
    
    # Fetch orders
    try:
        response = requests.get(
            orders_endpoint,
            params={
                "after": start_iso,
                "before": end_iso,
                "per_page": 100  # adjust as needed
            },
            auth=(WC_CONSUMER_KEY, WC_CONSUMER_SECRET)
        )
        response.raise_for_status()
        orders = response.json()
        st.success(f"Fetched {len(orders)} orders.")

        # ------------------------
        # Transform orders into CSV rows
        csv_rows = []
        for order in orders:
            order_id = order["id"]
            invoice_number = f"INV-{datetime.now().strftime('%Y%m%d')}-{order_id}"
            invoice_date = datetime.strptime(order["date_created"], "%Y-%m-%dT%H:%M:%S").strftime("%Y-%m-%d %H:%M:%S")
            invoice_status = order["status"].capitalize()
            customer_name = f"{order['billing']['first_name']} {order['billing']['last_name']}"
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
                    "Invoice Status": invoice_status,
                    "Customer Name": customer_name,
                    "Place of Supply": place_of_supply,
                    "Currency Code": currency,
                    "Item Name": item["name"],
                    "HSN/SAC": hsn,
                    "Item Type": item.get("type", "goods"),
                    "Quantity": item["quantity"],
                    "Usage unit": usage_unit,
                    "Item Price": item["price"],
                    "Is Inclusive Tax": "FALSE",  # simplify for now
                    "Item Tax %": item.get("tax_class", "0"),
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
        st.dataframe(df.head(10))  # preview first 10 rows

        # ------------------------
        # CSV download
        csv_bytes = df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="Download CSV",
            data=csv_bytes,
            file_name=f"orders_{start_date}_{end_date}.csv",
            mime="text/csv"
        )

    except requests.exceptions.RequestException as e:
        st.error(f"Error fetching orders: {e}")
