import streamlit as st
import pandas as pd
import requests
from io import BytesIO
from zipfile import ZipFile
from datetime import datetime
import os

# === WOO API CREDENTIALS ===
WC_API_URL = st.secrets.get("WC_API_URL")
WC_CONSUMER_KEY = st.secrets.get("WC_CONSUMER_KEY")
WC_CONSUMER_SECRET = st.secrets.get("WC_CONSUMER_SECRET")

if not WC_API_URL or not WC_CONSUMER_KEY or not WC_CONSUMER_SECRET:
    st.error("WooCommerce API credentials are missing. Please add them to Streamlit secrets.")
    st.stop()

# === FUNCTIONS ===
def fetch_orders(start_date, end_date):
    """Fetch WooCommerce orders between start_date and end_date"""
    url = f"{WC_API_URL}/orders"
    params = {
        "consumer_key": WC_CONSUMER_KEY,
        "consumer_secret": WC_CONSUMER_SECRET,
        "after": f"{start_date}T00:00:00",
        "before": f"{end_date}T23:59:59",
        "per_page": 100
    }
    response = requests.get(url, params=params)
    response.raise_for_status()
    return response.json()

def load_item_database():
    """Load item_database.xlsx from the app folder"""
    files_in_dir = [f.lower() for f in os.listdir(".")]
    if "item_database.xlsx" not in files_in_dir:
        st.error("item_database.xlsx not found in the app directory.")
        st.stop()

    df = pd.read_excel("item_database.xlsx")

    # Validate required columns
    required_cols = {"no", "woocommerce name", "zoho name", "hsn", "usage count"}
    df_cols_lower = set(col.lower() for col in df.columns)
    if not required_cols.issubset(df_cols_lower):
        st.error(f"item_database.xlsx is missing required columns. Expected: {required_cols}")
        st.stop()

    st.write("Item database loaded successfully. Preview below:")
    st.dataframe(df.head())  # show first 5 rows to confirm
    return df

def create_summary_and_csv(order_data, item_db):
    """Create the order DataFrame and summary based on Woo data and item database mapping"""

    records = []
    for order in order_data:
        order_number = order["number"]
        order_date = order["date_created"]
        customer_name = order["billing"]["first_name"] + " " + order["billing"]["last_name"]
        order_total = float(order["total"])

        # Loop through line items
        for item in order["line_items"]:
            wc_item_name = item["name"]

            # Match with item database (case-insensitive exact match)
            match = item_db[item_db["woocommerce name"].str.lower() == wc_item_name.lower()]
            if not match.empty:
                zoho_name = match.iloc[0]["zoho name"]
                hsn = str(match.iloc[0]["hsn"])  # Keep leading zeros
                usage_count = match.iloc[0]["usage count"]
            else:
                zoho_name = wc_item_name
                hsn = ""
                usage_count = ""

            records.append({
                "Invoice Number": "",  # Placeholder
                "Order Number": order_number,
                "Date": order_date,
                "Customer Name": customer_name,
                "Item Name": zoho_name,
                "HSN": hsn,
                "Usage Count": usage_count,
                "Order Total": order_total
            })

    df = pd.DataFrame(records)

    # Create summary table
    summary = df.groupby("Order Number").agg({
        "Order Total": "first",
        "Customer Name": "first"
    }).reset_index()

    # Add grand total
    grand_total = df["Order Total"].sum()
    summary.loc[len(summary)] = ["Grand Total", "", grand_total]

    return df, summary

def create_excel(summary_df):
    """Generate Excel file with nice formatting"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        summary_df.to_excel(writer, index=False, sheet_name="Summary")

        # Auto-adjust column width
        workbook = writer.book
        worksheet = writer.sheets["Summary"]
        for idx, col in enumerate(summary_df.columns):
            max_len = max(summary_df[col].astype(str).map(len).max(), len(col))
            worksheet.set_column(idx, idx, max_len + 2)
    output.seek(0)
    return output

# === STREAMLIT APP ===
st.title("WooCommerce to Zoho Reconciliation")

# Date inputs
start_date = st.date_input("Start Date", value=datetime.now().date())
end_date = st.date_input("End Date", value=datetime.now().date())

if st.button("Fetch Orders"):
    with st.spinner("Fetching data from WooCommerce..."):
        try:
            orders = fetch_orders(start_date, end_date)
            st.success(f"Fetched {len(orders)} orders from WooCommerce.")

            # Load item database
            item_db = load_item_database()

            # Create summary and main df
            df, summary_df = create_summary_and_csv(orders, item_db)

            st.subheader("Summary Preview")
            st.dataframe(summary_df.head())

            # Generate Excel
            excel_data = create_excel(summary_df)

            # CSV Download
            csv_data = df.to_csv(index=False).encode('utf-8')

            # === DOWNLOAD BUTTONS ===
            st.download_button(
                label="Download Orders CSV",
                data=csv_data,
                file_name=f"orders_{start_date}_to_{end_date}.csv",
                mime="text/csv"
            )

            st.download_button(
                label="Download Summary Excel",
                data=excel_data,
                file_name=f"summary_{start_date}_to_{end_date}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error fetching data: {e}")
