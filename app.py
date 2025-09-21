import streamlit as st
import pandas as pd
import requests
from io import BytesIO
from datetime import datetime

# ==============================
# WooCommerce API credentials
# ==============================
WC_API_URL = st.secrets.get("WC_API_URL")
WC_CONSUMER_KEY = st.secrets.get("WC_CONSUMER_KEY")
WC_CONSUMER_SECRET = st.secrets.get("WC_CONSUMER_SECRET")

if not WC_API_URL or not WC_CONSUMER_KEY or not WC_CONSUMER_SECRET:
    st.error("WooCommerce API credentials are missing. Please add them to Streamlit secrets.")
    st.stop()

st.title("WooCommerce → Zoho Reconciliation")

# ==============================
# Load Item Database
# ==============================
try:
    # Read item_database.xlsx and normalize headers
    item_db_df = pd.read_excel("item_database.xlsx", dtype={"HSN": str})
    item_db_df.columns = [col.strip().lower() for col in item_db_df.columns]

    # Validate required columns
    required_columns = ["woocommerce name", "zoho name", "hsn", "usage unit"]
    for col in required_columns:
        if col not in item_db_df.columns:
            st.error(f"Missing required column in item_database.xlsx: '{col}'")
            st.stop()

    st.subheader("Item Database")
    st.dataframe(item_db_df)

    # Create mapping for WooCommerce -> Zoho name
    name_mapping = {
        str(row["woocommerce name"]).lower(): str(row["zoho name"])
        for _, row in item_db_df.iterrows()
    }

    # Also keep direct dataframe lookup for HSN and Usage unit
    item_db_lookup = item_db_df.set_index("zoho name")[["hsn", "usage unit"]]

except FileNotFoundError:
    st.warning("item_database.xlsx not found. Please upload it to the app folder.")
    name_mapping = {}
    item_db_lookup = pd.DataFrame()
except Exception as e:
    st.error(f"Error reading item_database.xlsx: {e}")
    name_mapping = {}
    item_db_lookup = pd.DataFrame()

# ==============================
# Fetch WooCommerce Orders
# ==============================
st.subheader("Fetch WooCommerce Orders")
start_date = st.date_input("Start Date", value=datetime.today().replace(day=1))
end_date = st.date_input("End Date", value=datetime.today())

if st.button("Fetch Orders"):
    try:
        # WooCommerce API request
        url = f"{WC_API_URL}/orders"
        params = {
            "after": f"{start_date}T00:00:00",
            "before": f"{end_date}T23:59:59",
            "per_page": 100
        }
        response = requests.get(url, params=params, auth=(WC_CONSUMER_KEY, WC_CONSUMER_SECRET))
        response.raise_for_status()

        orders = response.json()

        if not orders:
            st.warning("No orders found for the given date range.")
            st.stop()

        # ==============================
        # Process Orders
        # ==============================
        order_data = []
        replaced_log = []  # log replacements for user visibility

        for order in orders:
            order_number = order.get("number")
            order_total = order.get("total")
            order_status = order.get("status")
            order_date = order.get("date_created")

            for item in order.get("line_items", []):
                original_name = item.get("name", "").strip()
                item_name_lower = original_name.lower()

                # Replace WooCommerce name with Zoho name if match found
                if item_name_lower in name_mapping:
                    zoho_name = name_mapping[item_name_lower]

                    # Fetch HSN and Usage Unit
                    hsn_value = item_db_lookup.loc[zoho_name, "hsn"] if zoho_name in item_db_lookup.index else ""
                    usage_unit_value = item_db_lookup.loc[zoho_name, "usage unit"] if zoho_name in item_db_lookup.index else ""

                    replaced_log.append({
                        "original_name": original_name,
                        "replaced_with": zoho_name,
                        "hsn": hsn_value,
                        "usage_unit": usage_unit_value
                    })
                else:
                    zoho_name = original_name
                    hsn_value = ""
                    usage_unit_value = ""

                order_data.append({
                    "Invoice Number": "",  # placeholder for Zoho invoice
                    "Order Number": order_number,
                    "Date": order_date,
                    "Name": zoho_name,
                    "Order Total": order_total,
                    "HSN": hsn_value,
                    "Usage Unit": usage_unit_value,
                    "Status": order_status
                })

        df_orders = pd.DataFrame(order_data)

        # ==============================
        # Summary
        # ==============================
        st.subheader("Summary")
        status_counts = df_orders["Status"].value_counts().to_dict()
        st.write("Order Status Counts:", status_counts)
        st.write("Total Orders:", len(df_orders))

        # ==============================
        # Show Replacement Log
        # ==============================
        if replaced_log:
            st.subheader("Replaced Item Names Log")
            st.dataframe(pd.DataFrame(replaced_log))
        else:
            st.info("No WooCommerce names were replaced with Zoho names.")

        # ==============================
        # Export as CSV
        # ==============================
        output = BytesIO()
        df_orders.to_csv(output, index=False)
        st.download_button(
            label="Download Reconciliation CSV",
            data=output.getvalue(),
            file_name=f"reconciliation_{start_date}_to_{end_date}.csv",
            mime="text/csv"
        )

    except requests.exceptions.HTTPError as e:
        st.error(f"HTTP Error: {e}")
    except Exception as e:
        st.error(f"Error fetching orders: {e}")
