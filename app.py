import streamlit as st
from pptx import Presentation
from pptx.table import Table
from pptx.shapes.group import GroupShape
from pptx.shapes.picture import Picture
from io import BytesIO
from balaram_converter import convert_balaram_to_unicode

# --- Streamlit Page Config ---
st.set_page_config(page_title="Balaram to Unicode", layout="centered")

# --- Load Custom CSS ---
with open("style.css") as f:
    st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

# --- Title Only (No śloka or subtitle) ---
st.markdown("<h1>📘 Balaram to Unicode PPTX Converter</h1>", unsafe_allow_html=True)

st.divider()

# --- File Upload ---
uploaded_file = st.file_uploader("📂 Upload your .pptx file", type=["pptx"])

# --- Converter Logic ---
def convert_text_frame(tf):
    if tf:
        for para in tf.paragraphs:
            for run in para.runs:
                run.text = convert_balaram_to_unicode(run.text)

def convert_table(table: Table):
    for row in table.rows:
        for cell in row.cells:
            convert_text_frame(cell.text_frame)

def process_shape(shape):
    if isinstance(shape, Picture): return
    if shape.has_text_frame:
        convert_text_frame(shape.text_frame)
    elif shape.shape_type == 19:  # Table
        convert_table(shape.table)
    elif isinstance(shape, GroupShape):
        for subshape in shape.shapes:
            process_shape(subshape)

def convert_pptx(pptx_file):
    prs = Presentation(pptx_file)
    for slide in prs.slides:
        for shape in slide.shapes:
            process_shape(shape)
    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return output

# --- Conversion Trigger ---
if uploaded_file:
    with st.spinner("🔄 Converting..."):
        result = convert_pptx(uploaded_file)
        st.success("✅ Conversion complete!")
        st.download_button(
            "📥 Download Converted PPTX",
            result,
            file_name="converted_balaram_unicode.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

# --- Footer ---
st.markdown("<hr>", unsafe_allow_html=True)
st.markdown("<div style='text-align:center; font-size:16px;'>🌸 Hare Kṛṣṇa! All glories to Śrīla Prabhupāda. 🌸</div>", unsafe_allow_html=True)
