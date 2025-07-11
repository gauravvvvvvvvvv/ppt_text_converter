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

LOCK_FILE = "/tmp/pptx_converter_user.lock"

def get_device_id():
    if "device_id" not in st.session_state:
        st.session_state.device_id = hashlib.md5(str(time.time()).encode()).hexdigest()[:10]
    return st.session_state.device_id

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

def is_locked():
    if os.path.exists(LOCK_FILE):
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

if is_locked():
    current_user = get_lock_user()
    if current_user != CURRENT_USER:
        st.warning(f"🚦 {current_user} is currently using the converter. Please wait.")
        st.stop()

st.set_page_config(page_title="Balaram to Unicode Converter", page_icon="📘", layout="centered")

def load_css():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Crimson+Text:ital,wght@0,400;0,600;1,400&display=swap');

    html, body {
        background: linear-gradient(135deg, #fffdf4 0%, #faf8f0 100%);
        font-family: 'Crimson Text', 'Georgia', serif;
        color: #4b2e0f;
        line-height: 1.6;
    }

    h1, h2, h3 {
        color: #6d3600;
        text-align: center;
        font-weight: 600;
        text-shadow: 0 1px 2px rgba(109, 54, 0, 0.1);
    }

    /* Enhanced styling for Balaram Unicode Converter */
@import url('https://fonts.googleapis.com/css2?family=Crimson+Text:ital,wght@0,400;0,600;1,400&display=swap');

html, body {
    background: linear-gradient(135deg, #fffdf4 0%, #faf8f0 100%);
    font-family: 'Crimson Text', 'Georgia', serif;
    color: #4b2e0f;
    line-height: 1.6;
}

/* Header styling */
h1, h2, h3 {
    color: #6d3600;
    text-align: center;
    font-weight: 600;
    text-shadow: 0 1px 2px rgba(109, 54, 0, 0.1);
}

h1 {
    font-size: 2.5rem;
    margin-bottom: 0.5rem;
}

/* Button styling with enhanced effects */
.stButton>button, .stDownloadButton>button {
    background: linear-gradient(135deg, #b06e11 0%, #8f5409 100%) !important;
    color: white !important;
    font-weight: bold;
    font-family: 'Crimson Text', serif;
    border-radius: 12px !important;
    padding: 12px 24px !important;
    border: none !important;
    transition: all 0.3s ease !important;
    box-shadow: 0 4px 15px rgba(176, 110, 17, 0.3);
    text-transform: none !important;
    font-size: 1.1rem !important;
}

.stButton>button:hover, .stDownloadButton>button:hover {
    background: linear-gradient(135deg, #8f5409 0%, #6d3600 100%) !important;
    transform: translateY(-2px);
    box-shadow: 0 6px 20px rgba(176, 110, 17, 0.4);
}

.stButton>button:active, .stDownloadButton>button:active {
    transform: translateY(0px);
}

/* File uploader enhanced styling */
div[data-testid="stFileUploader"] {
    background: linear-gradient(135deg, #fff5dc 0%, #ffeaa7 100%);
    border: 3px dashed #e0a958;
    padding: 30px;
    border-radius: 16px;
    text-align: center;
    transition: all 0.3s ease;
    position: relative;
    overflow: hidden;
}

div[data-testid="stFileUploader"]:hover {
    border-color: #b06e11;
    background: linear-gradient(135deg, #fff5dc 0%, #f9e79f 100%);
    transform: scale(1.02);
}

div[data-testid="stFileUploader"]::before {
    content: '';
    position: absolute;
    top: -2px;
    left: -2px;
    right: -2px;
    bottom: -2px;
    background: linear-gradient(45deg, #e0a958, #b06e11, #e0a958);
    border-radius: 16px;
    z-index: -1;
    opacity: 0;
    transition: opacity 0.3s ease;
}

div[data-testid="stFileUploader"]:hover::before {
    opacity: 0.3;
}

div[data-testid="stFileUploader"] > div {
    background-color: transparent !important;
}

/* Spinner styling */
.stSpinner > div {
    border-top-color: #b06e11 !important;
}

/* Enhanced message styling */
.stSuccess {
    background: linear-gradient(135deg, #e8f5e8 0%, #d4edda 100%);
    border-left: 5px solid #28a745;
    border-radius: 8px;
    padding: 1rem;
    box-shadow: 0 2px 10px rgba(40, 167, 69, 0.1);
}

.stInfo {
    background: linear-gradient(135deg, #e3f2fd 0%, #cce7ff 100%);
    border-left: 5px solid #2196f3;
    border-radius: 8px;
    padding: 1rem;
    box-shadow: 0 2px 10px rgba(33, 150, 243, 0.1);
}

.stWarning {
    background: linear-gradient(135deg, #fff3e0 0%, #ffe0b3 100%);
    border-left: 5px solid #ff9800;
    border-radius: 8px;
    padding: 1rem;
    box-shadow: 0 2px 10px rgba(255, 152, 0, 0.1);
}

.stError {
    background: linear-gradient(135deg, #ffebee 0%, #ffcdd2 100%);
    border-left: 5px solid #f44336;
    border-radius: 8px;
    padding: 1rem;
    box-shadow: 0 2px 10px rgba(244, 67, 54, 0.1);
}

/* Hide Streamlit branding */
footer {
    visibility: hidden;
}

#MainMenu {
    visibility: hidden;
}

header[data-testid="stHeader"] {
    display: none;
}

/* Custom divider */
hr {
    border: none;
    height: 3px;
    background: linear-gradient(90deg, transparent, #e0a958, #b06e11, #e0a958, transparent);
    margin: 30px 0;
    border-radius: 2px;
}

/* Expander styling */
.streamlit-expanderHeader {
    background: linear-gradient(135deg, #fff5dc 0%, #ffeaa7 100%);
    border-radius: 12px;
    color: #6d3600 !important;
    font-weight: bold;
    font-family: 'Crimson Text', serif;
    padding: 0.75rem 1rem;
    border: 2px solid #e0a958;
    transition: all 0.3s ease;
}

.streamlit-expanderHeader:hover {
    background: linear-gradient(135deg, #ffeaa7 0%, #f9e79f 100%);
    border-color: #b06e11;
}

.streamlit-expanderContent {
    background-color: rgba(255, 245, 220, 0.3);
    border-radius: 0 0 12px 12px;
    padding: 1rem;
    border: 2px solid #e0a958;
    border-top: none;
}

/* Code blocks */
.stCode {
    background-color: #f8f9fa !important;
    border: 1px solid #e0a958 !important;
    border-radius: 8px !important;
    font-family: 'Courier New', monospace !important;
}

/* Columns for better layout */
.stColumn {
    padding: 0 0.5rem;
}

/* Progress styling */
.stProgress .st-bo {
    background-color: #e0a958 !important;
}

/* Custom animations */
@keyframes fadeIn {
    from { opacity: 0; transform: translateY(20px); }
    to { opacity: 1; transform: translateY(0); }
}

.main .block-container {
    animation: fadeIn 0.6s ease-out;
}

/* Responsive design */
@media (max-width: 768px) {
    h1 {
        font-size: 2rem;
    }
    
    div[data-testid="stFileUploader"] {
        padding: 20px;
    }
    
    .stButton>button, .stDownloadButton>button {
        padding: 10px 20px !important;
        font-size: 1rem !important;
    }
}

    </style>
    """, unsafe_allow_html=True)


load_css()

st.markdown("<h1>📘 Balaram to Unicode PPTX Converter</h1>", unsafe_allow_html=True)
st.markdown(f"<p style='text-align: center; color: #6d3600;'>Welcome, <b>{CURRENT_USER}</b>! You're the active user.</p>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("📂 Upload your .pptx file", type=["pptx"])
just_unlock = st.checkbox("🔓 Only unlock file (no conversion)", value=False)

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
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp_input:
            tmp_input.write(pptx_bytes)
            tmp_input.flush()
            tmp_path = tmp_input.name

        prs = Presentation(tmp_path)
        for slide in prs.slides:
            for shape in slide.shapes:
                process_shape(shape)

        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp_output:
            prs.save(tmp_output.name)
            tmp_output.seek(0)
            with open(tmp_output.name, 'rb') as f:
                return BytesIO(f.read())

    except Exception as e:
        print(f"❌ Error during conversion: {e}")
        return None

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

with st.expander("ℹ️ How to use this converter"):
    st.markdown("""
    1. Enter your name  
    2. Upload a `.pptx` file  
    3. Choose whether to unlock only or convert  
    4. Download the result  
    5. Only one user can process at a time  
    """)

st.markdown(
    "<div style='text-align:center; font-size:16px; margin-top: 20px; color: #6d3600;'>"
    "🌸 Hare Kṛṣṇa! All glories to Śrīla Prabhupāda. 🌸</div>",
    unsafe_allow_html=True
)
