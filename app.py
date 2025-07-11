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
import time
import hashlib

# Path for lock and user tracking
LOCK_FILE = "/tmp/pptx_converter_user.lock"

# Helper to get unique device ID (stored locally via Streamlit session)
def get_device_id():
    if "device_id" not in st.session_state:
        st.session_state.device_id = hashlib.md5(str(time.time()).encode()).hexdigest()[:10]
    return st.session_state.device_id

# Ask name if not set
if "user_name" not in st.session_state:
    st.session_state.user_name = None

def ask_for_name():
    st.title("🧑 Who's using the converter?")
    name = st.text_input("Enter your name (no password needed):")
    if name.strip():
        st.session_state.user_name = name.strip()
        st.rerun()

if not st.session_state.user_name:
    ask_for_name()
    st.stop()

CURRENT_USER = st.session_state.user_name
DEVICE_ID = get_device_id()

# Locking logic
def is_locked():
    if os.path.exists(LOCK_FILE):
        # Clear stale lock after 5 min
        if time.time() - os.path.getmtime(LOCK_FILE) > 300:
            os.remove(LOCK_FILE)
            return False
        return True
    return False

def get_lock_user():
    try:
        with open(LOCK_FILE, 'r') as f:
            return f.read().strip()
    except:
        return None

def acquire_lock(user):
    with open(LOCK_FILE, "w") as f:
        f.write(user)

def release_lock():
    if os.path.exists(LOCK_FILE):
        os.remove(LOCK_FILE)

# Lock gatekeeping
if is_locked():
    current_user = get_lock_user()
    if current_user != CURRENT_USER:
        st.warning(f"🚦 {current_user} is currently using the converter.\n\nPlease wait until they finish.")
        st.stop()

# ========== UI and Setup ==========
st.set_page_config(page_title="Balaram to Unicode Converter", page_icon="📘", layout="centered")

def load_css():
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

st.markdown("<h1>📘 Balaram to Unicode PPTX Converter</h1>", unsafe_allow_html=True)
st.markdown(f"<p style='text-align: center; color: #6d3600;'>Welcome, <b>{CURRENT_USER}</b>! You're the active user.</p>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("📂 Upload your .pptx file", type=["pptx"])
just_unlock = st.checkbox("🔓 Only unlock file (no conversion)", value=False)

# ========== Conversion Logic ==========
balaram_map = {
    'ä': 'ā', 'é': 'ī', 'ü': 'ū', 'å': 'ṛ', 'è': 'ṝ', 'ì': 'ṅ', 'ï': 'ñ', 'ö': 'ṭ',
    'ò': 'ḍ', 'ë': 'ṇ', 'ç': 'ś', 'à': 'ṁ', 'ù': 'ḥ', 'ÿ': 'ḷ', 'û': 'ḹ', 'ý': 'ẏ',
    'Ä': 'Ā', 'É': 'Ī', 'Ü': 'Ū', 'Å': 'Ṛ', 'È': 'Ṝ', 'Ì': 'Ṅ', 'Ï': 'Ñ',
    'Ö': 'Ṭ', 'Ò': 'Ḍ', 'Ë': 'Ṇ', 'Ç': 'Ś', 'À': 'Ṁ', 'Ù': 'Ḥ', 'ß': 'Ḷ',
    'Ý': 'Ẏ', '~': 'ɱ', "'": "'", '…': '…', '’': '’', 'ñ': 'ṣ', 'Ñ': 'Ṣ'
}

def convert_balaram_to_unicode(text: str) -> str:
    return ''.join(balaram_map.get(char, char) for char in text)

def convert_text_frame(tf):
    if tf and tf.text.strip():
        for para in tf.paragraphs:
            for run in para.runs:
                run.text = convert_balaram_to_unicode(run.text)
        return True
    return False

def convert_table(table: Table):
    conversions = 0
    for row in table.rows:
        for cell in row.cells:
            if convert_text_frame(cell.text_frame):
                conversions += 1
    return conversions

def process_shape(shape):
    conversions = 0
    if isinstance(shape, Picture): return 0
    try:
        if shape.has_text_frame:
            if convert_text_frame(shape.text_frame):
                conversions += 1
        elif hasattr(shape, 'shape_type') and shape.shape_type == 19:
            conversions += convert_table(shape.table)
        elif isinstance(shape, GroupShape):
            for subshape in shape.shapes:
                conversions += process_shape(subshape)
    except:
        pass
    return conversions

def unlock_pptx_file(pptx_bytes, filename):
    with tempfile.TemporaryDirectory() as temp_dir:
        zip_temp = os.path.join(temp_dir, "temp.zip")
        extract_path = os.path.join(temp_dir, "extract")
        with open(zip_temp, 'wb') as f:
            f.write(pptx_bytes)
        try:
            with zipfile.ZipFile(zip_temp, 'r') as zip_ref:
                zip_ref.extractall(extract_path)
        except:
            return pptx_bytes
        pres_xml = os.path.join(extract_path, 'ppt', 'presentation.xml')
        if os.path.exists(pres_xml):
            try:
                ET.register_namespace('p', "http://schemas.openxmlformats.org/presentationml/2006/main")
                tree = ET.parse(pres_xml)
                root = tree.getroot()
                for elem in root.findall('{http://schemas.openxmlformats.org/presentationml/2006/main}modifyVerifier'):
                    root.remove(elem)
                for elem in root.findall('.//modifyVerifier'):
                    try:
                        root.remove(elem)
                    except:
                        pass
                tree.write(pres_xml, encoding='utf-8', xml_declaration=True)
            except:
                pass
        try:
            output_zip = os.path.join(temp_dir, "unlocked.zip")
            shutil.make_archive(output_zip.replace('.zip', ''), 'zip', extract_path)
            with open(output_zip, 'rb') as f:
                return f.read()
        except:
            return pptx_bytes

def convert_pptx(pptx_bytes):
    try:
        prs = Presentation(BytesIO(pptx_bytes))
        for slide in prs.slides:
            for shape in slide.shapes:
                process_shape(shape)
        output = BytesIO()
        prs.save(output)
        output.seek(0)
        return output
    except:
        return None

# ========== Processing ==========
if uploaded_file is not None:
    try:
        acquire_lock(CURRENT_USER)

        file_bytes = uploaded_file.getvalue()
        st.write(f"📄 File size: `{len(file_bytes) / 1024**2:.2f} MB`")
        unlocked_bytes = unlock_pptx_file(file_bytes, uploaded_file.name)
        st.write(f"🔓 Unlocked size: `{len(unlocked_bytes) / 1024**2:.2f} MB`")

        if just_unlock:
            st.download_button(
                label="📥 Download Unlocked PPTX",
                data=unlocked_bytes,
                file_name=f"{os.path.splitext(uploaded_file.name)[0]}_unlocked.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True
            )
        else:
            converted_stream = convert_pptx(unlocked_bytes)
            if converted_stream:
                st.download_button(
                    label="📥 Download Converted PPTX",
                    data=converted_stream,
                    file_name=f"{os.path.splitext(uploaded_file.name)[0]}_unicode.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True
                )
            else:
                st.error("❌ Conversion failed.")
    except Exception as e:
        st.error(f"❌ Error: {e}")
    finally:
        release_lock()

# ========== Info ==========
with st.expander("ℹ️ How to use this converter"):
    st.markdown("""
    1. Enter your name (first time only)  
    2. Upload `.pptx` file  
    3. Wait if someone else is using it  
    4. Choose unlock or convert  
    5. Download result  
    """)

st.markdown(
    "<div style='text-align:center; font-size:16px; margin-top: 20px; color: #6d3600;'>"
    "🌸 Hare Kṛṣṇa! All glories to Śrīla Prabhupāda. 🌸</div>",
    unsafe_allow_html=True
)
