import streamlit as st
import requests
import pandas as pd
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.pdfbase.pdfmetrics import stringWidth
from datetime import datetime

# ------------------------
# WooCommerce API settings
WC_API_URL = st.secrets.get("WC_API_URL")
WC_CONSUMER_KEY = st.secrets.get("WC_CONSUMER_KEY")
WC_CONSUMER_SECRET = st.secrets.get("WC_CONSUMER_SECRET")

if not WC_API_URL or not WC_CONSUMER_KEY or not WC_CONSUMER_SECRET:
    st.error("WooCommerce API credentials are missing. Please add them to Streamlit secrets.")
    st.stop()

# -----------------------------
# STREAMLIT APP CONFIG
# -----------------------------
st.set_page_config(page_title="WooCommerce Invoice Generator", layout="wide")

# -----------------------------
# USER INPUTS
# -----------------------------
st.title("WooCommerce Invoice Generator")


col1, col2 = st.columns(2)
with col1:
    start_date = st.date_input("Start Date")
with col2:
    end_date = st.date_input("End Date")

invoice_prefix = st.text_input("Invoice Prefix", value="ECHE/2526/")
start_sequence = st.number_input("Starting Sequence Number", min_value=1, value=608)

# -----------------------------
# HELPER FUNCTIONS
# -----------------------------
def fetch_orders():
    """Fetch orders from WooCommerce within date range"""
    page = 1
    per_page = 100
    orders = []
    while True:
        response = requests.get(
            f"{api_url}/wp-json/wc/v3/orders",
            params={
                "consumer_key": consumer_key,
                "consumer_secret": consumer_secret,
                "after": f"{start_date}T00:00:00",
                "before": f"{end_date}T23:59:59",
                "per_page": per_page,
                "page": page,
            },
        )
        if response.status_code != 200:
            st.error(f"Failed to fetch orders: {response.text}")
            return []

        data = response.json()
        if not data:
            break
        orders.extend(data)
        page += 1

    return orders

def draw_invoice(c, invoice_number, order_number, customer_name):
    """Draw a single invoice on the PDF"""
    y = 270  # Start position from top
    c.setFont("Courier-Bold", 12)
    c.drawString(20*mm, y*mm, f"Invoice: {invoice_number}")
    y -= 10
    c.setFont("Courier", 11)
    c.drawString(20*mm, y*mm, f"Order Number: {order_number}")
    y -= 10
    c.drawString(20*mm, y*mm, f"Customer Name: {customer_name}")
    y -= 10
    c.line(20*mm, y*mm, 190*mm, y*mm)  # Horizontal separator line

# -----------------------------
# MAIN PROCESS
# -----------------------------
if st.button("Generate Invoices"):
    if not (api_url and consumer_key and consumer_secret and start_date and end_date):
        st.warning("Please fill in all fields before generating invoices.")
    else:
        st.info("Fetching orders, please wait...")
        all_orders = fetch_orders()

        if not all_orders:
            st.warning("No orders found for the selected date range.")
        else:
            # Convert WooCommerce data into CSV rows
            csv_rows = []
            sequence_number = start_sequence
            total_revenue = 0
            status_counts = {
                "completed": 0,
                "processing": 0,
                "on-hold": 0,
                "cancelled": 0,
                "pending": 0,
            }

            for order in all_orders:
                # Count statuses
                status = order.get("status", "").lower()
                if status in status_counts:
                    status_counts[status] += 1

                # Get WooCommerce total
                order_total = float(order.get("total", 0.0))
                total_revenue += order_total

                # Build reconciliation CSV row
                invoice_number = f"{invoice_prefix}{sequence_number:05d}"
                order_date = datetime.strptime(order["date_created"], "%Y-%m-%dT%H:%M:%S").strftime("%d-%m-%Y")
                billing_name = f"{order['billing']['first_name']} {order['billing']['last_name']}".strip()

                csv_rows.append({
                    "Invoice Number": invoice_number,
                    "Order Number": order["id"],
                    "Date": order_date,
                    "Name": billing_name,
                    "Order Total": order_total
                })

                sequence_number += 1

            # Create dataframe for reconciliation
            df_reconciliation = pd.DataFrame(csv_rows)

            # -----------------------------
            # Generate PDF
            # -----------------------------
            buffer = BytesIO()
            c = canvas.Canvas(buffer, pagesize=(210*mm, 297*mm))  # A4

            sequence_number = start_sequence
            for order in all_orders:
                invoice_number = f"{invoice_prefix}{sequence_number:05d}"
                customer_name = f"{order['billing']['first_name']} {order['billing']['last_name']}".strip()
                order_number = order["id"]

                draw_invoice(c, invoice_number, order_number, customer_name)
                c.showPage()
                sequence_number += 1

            c.save()
            buffer.seek(0)

            # -----------------------------
            # DOWNLOAD BUTTONS
            # -----------------------------
            st.download_button(
                "Download Invoices PDF",
                data=buffer,
                file_name="invoices.pdf",
                mime="application/pdf"
            )

            st.download_button(
                "Download Reconciliation CSV",
                data=df_reconciliation.to_csv(index=False),
                file_name="reconciliation.csv",
                mime="text/csv"
            )

            # -----------------------------
            # SUMMARY REPORT
            # -----------------------------
            first_order_id = all_orders[0]["id"]
            last_order_id = all_orders[-1]["id"]
            first_invoice_number = f"{invoice_prefix}{start_sequence:05d}"
            last_invoice_number = f"{invoice_prefix}{sequence_number - 1:05d}"

            with st.expander("View Summary Report"):
                st.subheader("Summary Report")
                st.write(f"**Total Orders Processed:** {len(all_orders)}")
                st.write(f"**Order IDs:** {first_order_id} → {last_order_id}")
                st.write(f"**Invoice Numbers:** {first_invoice_number} → {last_invoice_number}")
                st.write(f"**Total Revenue (WooCommerce):** ₹ {total_revenue:,.2f}")

                st.markdown("### Order Status Breakdown")
                st.write(f"✅ Completed: {status_counts['completed']}")
                st.write(f"⚙️ Processing: {status_counts['processing']}")
                st.write(f"⏸️ On-Hold: {status_counts['on-hold']}")
                st.write(f"❌ Cancelled: {status_counts['cancelled']}")
                st.write(f"💰 Pending Payment: {status_counts['pending']}")
