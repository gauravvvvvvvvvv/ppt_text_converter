import streamlit as st
from pptx import Presentation
from pptx.table import Table
from pptx.shapes.group import GroupShape
from pptx.shapes.picture import Picture
from io import BytesIO
import os
import zipfile
import shutil
import xml.etree.ElementTree as ET
import tempfile
from balaram_converter import convert_balaram_to_unicode

# Page configuration
st.set_page_config(
    page_title="Balaram to Unicode Converter",
    page_icon="📘",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Load CSS styling
def load_css():
    try:
        with open("style.css") as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
    except FileNotFoundError:
        st.warning("CSS file not found. Using default styling.")

load_css()

# Header
st.markdown("<h1>📘 Balaram to Unicode PPTX Converter</h1>", unsafe_allow_html=True)
st.divider()

# File uploader
uploaded_file = st.file_uploader(
    "📂 Upload your .pptx file", 
    type=["pptx"],
    help="Select a PowerPoint file with Balaram font text to convert to Unicode"
)

def convert_text_frame(tf):
    """Convert text in a text frame from Balaram to Unicode"""
    if tf:
        for para in tf.paragraphs:
            for run in para.runs:
                if run.text:
                    run.text = convert_balaram_to_unicode(run.text)

def convert_table(table: Table):
    """Convert text in table cells from Balaram to Unicode"""
    for row in table.rows:
        for cell in row.cells:
            convert_text_frame(cell.text_frame)

def process_shape(shape):
    """Process different types of shapes in slides"""
    if isinstance(shape, Picture): 
        return
    
    if shape.has_text_frame:
        convert_text_frame(shape.text_frame)
    elif hasattr(shape, 'shape_type') and shape.shape_type == 19:  # Table
        convert_table(shape.table)
    elif isinstance(shape, GroupShape):
        for subshape in shape.shapes:
            process_shape(subshape)

def unlock_pptx_file(pptx_bytes, filename):
    """Remove protection from PPTX file by removing modifyVerifier elements"""
    with tempfile.TemporaryDirectory() as temp_dir:
        zip_temp = os.path.join(temp_dir, "temp.zip")
        extract_path = os.path.join(temp_dir, "extract")
        
        # Save uploaded file bytes to disk
        with open(zip_temp, 'wb') as f:
            f.write(pptx_bytes)
        
        # Extract the PPTX (which is a ZIP file)
        try:
            with zipfile.ZipFile(zip_temp, 'r') as zip_ref:
                zip_ref.extractall(extract_path)
        except Exception as e:
            st.error(f"Failed to extract PPTX: {e}")
            return pptx_bytes
        
        # Modify presentation.xml to remove protection
        pres_xml = os.path.join(extract_path, 'ppt', 'presentation.xml')
        unlocked = False
        
        if os.path.exists(pres_xml):
            try:
                # Register namespaces to preserve XML structure
                namespaces = {
                    '': "http://schemas.openxmlformats.org/presentationml/2006/main",
                    'a': "http://schemas.openxmlformats.org/drawingml/2006/main",
                    'r': "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                }
                
                for prefix, uri in namespaces.items():
                    ET.register_namespace(prefix, uri)
                
                tree = ET.parse(pres_xml)
                root = tree.getroot()
                
                # Remove modifyVerifier elements (protection)
                verifiers_removed = 0
                for elem in root.findall('.//{http://schemas.openxmlformats.org/presentationml/2006/main}modifyVerifier'):
                    root.remove(elem)
                    verifiers_removed += 1
                
                if verifiers_removed > 0:
                    tree.write(pres_xml, encoding='utf-8', xml_declaration=True)
                    unlocked = True
                    st.info(f"🔓 Removed {verifiers_removed} protection element(s)")
                else:
                    st.info("ℹ️ No protection found in this file")
                    
            except Exception as e:
                st.warning(f"⚠️ Could not modify protection in {filename}: {e}")
        
        # Repack into PPTX
        try:
            output_zip = os.path.join(temp_dir, "unlocked.zip")
            shutil.make_archive(output_zip.replace('.zip', ''), 'zip', extract_path)
            
            with open(output_zip, 'rb') as f:
                return f.read()
                
        except Exception as e:
            st.error(f"Failed to repack PPTX: {e}")
            return pptx_bytes

def convert_pptx(pptx_bytes):
    """Convert Balaram text to Unicode in PPTX"""
    try:
        prs = Presentation(BytesIO(pptx_bytes))
        
        slides_processed = 0
        shapes_processed = 0
        
        for slide in prs.slides:
            slides_processed += 1
            for shape in slide.shapes:
                process_shape(shape)
                shapes_processed += 1
        
        st.info(f"📊 Processed {slides_processed} slides with {shapes_processed} shapes")
        
        output = BytesIO()
        prs.save(output)
        output.seek(0)
        return output
        
    except Exception as e:
        st.error(f"Failed to convert PPTX: {e}")
        return None

# Main processing logic
if uploaded_file:
    # Read the uploaded file once
    file_bytes = uploaded_file.read()
    
    with st.spinner("🔓 Unlocking presentation (if protected)..."):
        unlocked_bytes = unlock_pptx_file(file_bytes, uploaded_file.name)
    
    with st.spinner("🔄 Converting Balaram text to Unicode..."):
        converted_stream = convert_pptx(unlocked_bytes)
        
        if converted_stream:
            original_name = os.path.splitext(uploaded_file.name)[0]
            converted_name = f"{original_name}_unicode.pptx"
            
            st.success("✅ Conversion completed successfully!")
            
            # Download button
            st.download_button(
                label="📥 Download Converted PPTX",
                data=converted_stream,
                file_name=converted_name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                help="Click to download your converted PowerPoint file"
            )
        else:
            st.error("❌ Conversion failed. Please check your file and try again.")

# Footer
st.markdown("<hr>", unsafe_allow_html=True)
st.markdown(
    "<div style='text-align:center; font-size:16px; margin-top: 20px;'>"
    "🌸 Hare Kṛṣṇa! All glories to Śrīla Prabhupāda. 🌸"
    "</div>", 
    unsafe_allow_html=True
)

# Instructions
with st.expander("ℹ️ How to use this converter"):
    st.markdown("""
    1. **Upload** your PowerPoint (.pptx) file using the file uploader above
    2. **Wait** for the conversion process to complete
    3. **Download** your converted file with Unicode text
    
    **Note**: This converter will:
    - Remove any presentation protection/locks
    - Convert Balaram font text to proper Unicode
    - Preserve all formatting and layouts
    """)
