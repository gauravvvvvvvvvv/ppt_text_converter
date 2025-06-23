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
        # Fallback inline CSS if file not found
        st.markdown("""
        <style>
        html, body { background-color: #fffdf4; font-family: 'Georgia', serif; color: #4b2e0f; }
        h1, h2, h3 { color: #6d3600; text-align: center; }
        .stButton>button, .stDownloadButton>button { 
            background-color: #b06e11 !important; color: white !important; 
            font-weight: bold; border-radius: 8px; padding: 10px 20px; 
        }
        div[data-testid="stFileUploader"] { 
            background-color: #fff5dc; border: 2px dashed #e0a958; 
            padding: 20px; border-radius: 12px; 
        }
        footer { visibility: hidden; }
        </style>
        """, unsafe_allow_html=True)

load_css()

# Header
st.markdown("<h1>📘 Balaram to Unicode PPTX Converter</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #6d3600; font-style: italic;'>Convert your PowerPoint presentations from Balaram font to Unicode</p>", unsafe_allow_html=True)
st.divider()

# File uploader
uploaded_file = st.file_uploader(
    "📂 Upload your .pptx file", 
    type=["pptx"],
    help="Select a PowerPoint file with Balaram font text to convert to Unicode"
)

def convert_text_frame(tf):
    """Convert text in a text frame from Balaram to Unicode"""
    if tf and tf.text.strip():  # Only process if there's actual text
        for para in tf.paragraphs:
            for run in para.runs:
                if run.text:
                    original_text = run.text
                    converted_text = convert_balaram_to_unicode(run.text)
                    run.text = converted_text
                    # Debug: Show conversion if text changed
                    if original_text != converted_text and len(original_text.strip()) > 0:
                        return True  # Indicate conversion happened
    return False

def convert_table(table: Table):
    """Convert text in table cells from Balaram to Unicode"""
    conversions = 0
    for row in table.rows:
        for cell in row.cells:
            if convert_text_frame(cell.text_frame):
                conversions += 1
    return conversions

def process_shape(shape):
    """Process different types of shapes in slides"""
    conversions = 0
    
    if isinstance(shape, Picture): 
        return 0
    
    try:
        if shape.has_text_frame:
            if convert_text_frame(shape.text_frame):
                conversions += 1
        elif hasattr(shape, 'shape_type') and shape.shape_type == 19:  # Table
            conversions += convert_table(shape.table)
        elif isinstance(shape, GroupShape):
            for subshape in shape.shapes:
                conversions += process_shape(subshape)
    except Exception as e:
        # Log shape processing errors but continue
        st.warning(f"⚠️ Could not process a shape: {str(e)[:100]}...")
    
    return conversions

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
            st.error(f"❌ Failed to extract PPTX: {e}")
            return pptx_bytes
        
        # Modify presentation.xml to remove protection
        pres_xml = os.path.join(extract_path, 'ppt', 'presentation.xml')
        
        if os.path.exists(pres_xml):
            try:
                # Register namespaces to preserve XML structure
                namespaces = {
                    'p': "http://schemas.openxmlformats.org/presentationml/2006/main",
                    'a': "http://schemas.openxmlformats.org/drawingml/2006/main",
                    'r': "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                }
                
                for prefix, uri in namespaces.items():
                    ET.register_namespace(prefix, uri)
                
                # Parse XML
                tree = ET.parse(pres_xml)
                root = tree.getroot()
                
                # Remove modifyVerifier elements (protection)
                verifiers_removed = 0
                
                # Try different XPath patterns to find modifyVerifier
                patterns = [
                    './/{http://schemas.openxmlformats.org/presentationml/2006/main}modifyVerifier',
                    './/p:modifyVerifier',
                    './/modifyVerifier'
                ]
                
                for pattern in patterns:
                    elements = root.findall(pattern)
                    for elem in elements:
                        parent = root.find(f'.//*[{elem.tag}]/..')
                        if parent is not None:
                            parent.remove(elem)
                            verifiers_removed += 1
                        else:
                            # Try direct removal from root
                            try:
                                root.remove(elem)
                                verifiers_removed += 1
                            except ValueError:
                                pass
                
                if verifiers_removed > 0:
                    # Write back with proper formatting
                    tree.write(pres_xml, encoding='utf-8', xml_declaration=True, method='xml')
                    st.success(f"🔓 Successfully removed {verifiers_removed} protection element(s)")
                else:
                    st.info("ℹ️ No protection elements found - file was not locked")
                    
            except Exception as e:
                st.warning(f"⚠️ Could not modify protection in {filename}: {e}")
                st.info("📝 Proceeding with conversion anyway...")
        
        # Repack into PPTX
        try:
            output_zip = os.path.join(temp_dir, "unlocked.zip")
            shutil.make_archive(output_zip.replace('.zip', ''), 'zip', extract_path)
            
            with open(output_zip, 'rb') as f:
                return f.read()
                
        except Exception as e:
            st.error(f"❌ Failed to repack PPTX: {e}")
            return pptx_bytes

def convert_pptx(pptx_bytes):
    """Convert Balaram text to Unicode in PPTX"""
    try:
        prs = Presentation(BytesIO(pptx_bytes))
        
        slides_processed = 0
        total_conversions = 0
        
        # Process each slide
        for slide_num, slide in enumerate(prs.slides, 1):
            slides_processed += 1
            slide_conversions = 0
            
            for shape in slide.shapes:
                slide_conversions += process_shape(shape)
            
            total_conversions += slide_conversions
            
            # Show progress for large presentations
            if slides_processed % 10 == 0:
                st.info(f"📊 Processed {slides_processed} slides so far...")
        
        # Show final statistics
        if total_conversions > 0:
            st.success(f"✨ Converted {total_conversions} text elements across {slides_processed} slides")
        else:
            st.warning("⚠️ No Balaram text found to convert. File processed anyway.")
        
        # Save converted presentation
        output = BytesIO()
        prs.save(output)
        output.seek(0)
        return output
        
    except Exception as e:
        st.error(f"❌ Failed to process PPTX: {e}")
        return None

# Main processing logic
if uploaded_file:
    # Show file info
    file_size = len(uploaded_file.read())
    uploaded_file.seek(0)  # Reset file pointer
    st.info(f"📄 File: {uploaded_file.name} ({file_size/1024:.1f} KB)")
    
    # Read the uploaded file
    file_bytes = uploaded_file.read()
    
    # Step 1: Unlock file
    with st.spinner("🔓 Removing presentation locks (if any)..."):
        unlocked_bytes = unlock_pptx_file(file_bytes, uploaded_file.name)
    
    # Step 2: Convert text
    with st.spinner("🔄 Converting Balaram text to Unicode..."):
        converted_stream = convert_pptx(unlocked_bytes)
        
        if converted_stream:
            original_name = os.path.splitext(uploaded_file.name)[0]
            converted_name = f"{original_name}_unicode.pptx"
            
            st.success("🎉 Conversion completed successfully!")
            
            # Download section
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                st.download_button(
                    label="📥 Download Converted PPTX",
                    data=converted_stream,
                    file_name=converted_name,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    help="Click to download your converted PowerPoint file",
                    use_container_width=True
                )
        else:
            st.error("❌ Conversion failed. Please check your file and try again.")

# Footer and instructions
st.markdown("<hr>", unsafe_allow_html=True)

# Instructions
with st.expander("ℹ️ How to use this converter"):
    st.markdown("""
    ### 📋 Instructions:
    1. **Upload** your PowerPoint (.pptx) file using the file uploader above
    2. **Wait** for the automatic processing (unlock + conversion)
    3. **Download** your converted file with proper Unicode text
    
    ### 🔧 What this tool does:
    - **Removes presentation protection/locks** automatically
    - **Converts Balaram font characters** to proper Unicode equivalents
    - **Preserves all formatting, layouts, and images**
    - **Processes all slides, shapes, tables, and text boxes**
    
    ### 📝 Supported conversions:
    The tool converts 40+ Balaram characters including:
    - `ä` → `ā` (long a)
    - `é` → `ī` (long i) 
    - `ì` → `ṅ` (nasal n)
    - `ç` → `ś` (palatal s)
    - And many more diacritical marks
    """)

# Sample conversion preview
with st.expander("🔤 Preview: Balaram to Unicode conversion"):
    sample_balarm = "Håñi Kåñëa Håñi Kåñëa, Kåñëa Kåñëa Håñi Håñi"
    sample_unicode = convert_balaram_to_unicode(sample_balarm)
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Balaram Font:**")
        st.code(sample_balarm, language=None)
    with col2:
        st.markdown("**Unicode Result:**")
        st.code(sample_unicode, language=None)

st.markdown(
    "<div style='text-align:center; font-size:16px; margin-top: 30px; color: #6d3600;'>"
    "🌸 Hare Kṛṣṇa! All glories to Śrīla Prabhupāda. 🌸"
    "</div>", 
    unsafe_allow_html=True
)
