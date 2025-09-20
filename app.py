import streamlit as st
from fpdf import FPDF
import os

# ==============================
# DEBUG - Check secrets
# ==============================
st.write("DEBUG SECRETS:", dict(st.secrets))

# ==============================
# Load WooCommerce API Keys
# ==============================
try:
    WC_CONSUMER_KEY = st.secrets["woocommerce"]["consumer_key"]
    WC_CONSUMER_SECRET = st.secrets["woocommerce"]["consumer_secret"]
    WC_STORE_URL = st.secrets["woocommerce"]["store_url"]
except KeyError as e:
    st.error(f"Missing WooCommerce secret key: {e}")
    st.stop()

# ==============================
# Custom PDF Class
# ==============================
class PDF(FPDF):
    def header(self):
        self.set_font("Roboto", "B", 16)
        self.cell(0, 10, "Generated Labels", 0, 1, "C")

    def footer(self):
        self.set_y(-15)
        self.set_font("Roboto", "I", 8)
        self.cell(0, 10, f"Page {self.page_no()}", 0, 0, "C")

# ==============================
# Load Roboto Font
# ==============================
font_path = os.path.join(os.path.dirname(__file__), "Roboto-Black.ttf")
if not os.path.exists(font_path):
    st.error(f"Font file not found: {font_path}")
    st.stop()

pdf = PDF()
pdf.add_font("Roboto", "", font_path, uni=True)
pdf.add_font("Roboto", "B", font_path, uni=True)
pdf.set_auto_page_break(auto=True, margin=15)
pdf.add_page()
pdf.set_font("Roboto", "B", 14)

# ==============================
# Example Content
# ==============================
pdf.cell(0, 10, "Hello, this is a test PDF using Roboto font!", ln=True)
pdf.ln(10)
pdf.set_font("Roboto", "", 12)
pdf.multi_cell(0, 10, "This PDF uses Roboto-Black.ttf instead of DejaVuSans.ttf.\n\nNo autodownload logic is required now.")

# ==============================
# Streamlit UI
# ==============================
st.title("PDF Generator")

# Generate PDF Button
if st.button("Generate PDF"):
    output_path = os.path.join(os.getcwd(), "labels.pdf")
    pdf.output(output_path)
    st.success(f"PDF generated: {output_path}")

    # Provide download link
    with open(output_path, "rb") as f:
        st.download_button("Download PDF", f, file_name="labels.pdf", mime="application/pdf")
