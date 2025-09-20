from fpdf import FPDF
import os
import io
from datetime import datetime
import streamlit as st

# ---- Streamlit App ----
st.title("WooCommerce to Zoho CSV Exporter")

# WooCommerce API Settings
WC_API_URL = "https://sustenance.co.in/wp-json/wc/v3"
WC_CONSUMER_KEY = st.secrets["woocommerce"]["consumer_key"]
WC_CONSUMER_SECRET = st.secrets["woocommerce"]["consumer_secret"]

# --- PDF Summary Function ---
class PDF(FPDF):
    def header(self):
        self.set_font("RobotoBlack", size=14)
        self.cell(0, 10, "WooCommerce Orders Summary", ln=True, align="C")
        self.ln(10)

def generate_summary_pdf(summary_data):
    pdf = PDF()
    pdf.add_page()

    # Add Roboto Black font
    pdf.add_font("RobotoBlack", "", "Roboto-Black.ttf", uni=True)
    pdf.set_font("RobotoBlack", size=12)

    # Add summary text
    pdf.cell(0, 10, f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", ln=True)
    pdf.ln(5)
    pdf.cell(0, 10, f"Number of Orders Downloaded: {summary_data['num_orders']}", ln=True)
    pdf.cell(0, 10, f"Order ID Range: {summary_data['start_order_id']} - {summary_data['end_order_id']}", ln=True)
    pdf.cell(0, 10, f"Invoice Number Range: {summary_data['start_invoice']} - {summary_data['end_invoice']}", ln=True)

    # Output PDF as bytes
    pdf_buffer = io.BytesIO()
    pdf.output(pdf_buffer)
    pdf_buffer.seek(0)
    return pdf_buffer

# --- Main Logic ---
if st.button("Generate CSV and Summary PDF"):
    # Example dummy summary data for now
    summary_data = {
        "num_orders": 15,
        "start_order_id": 101,
        "end_order_id": 115,
        "start_invoice": "INV-001",
        "end_invoice": "INV-015"
    }

    pdf_file = generate_summary_pdf(summary_data)

    # Allow PDF download
    st.download_button(
        label="Download Summary PDF",
        data=pdf_file,
        file_name=f"order_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
        mime="application/pdf"
    )
