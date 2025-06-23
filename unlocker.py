import os
import zipfile
import shutil
import xml.etree.ElementTree as ET
import tempfile

def unlock_pptx_file(pptx_bytes, filename):
    with tempfile.TemporaryDirectory() as temp_dir:
        zip_temp = os.path.join(temp_dir, "temp.zip")
        extract_path = os.path.join(temp_dir, "extract")

        # Save uploaded file to disk
        with open(zip_temp, 'wb') as f:
            f.write(pptx_bytes.read())

        with zipfile.ZipFile(zip_temp, 'r') as zip_ref:
            zip_ref.extractall(extract_path)

        pres_xml = os.path.join(extract_path, 'ppt', 'presentation.xml')
        if os.path.exists(pres_xml):
            try:
                ET.register_namespace('p', "http://schemas.openxmlformats.org/presentationml/2006/main")
                tree = ET.parse(pres_xml)
                root = tree.getroot()
                for elem in root.findall('{http://schemas.openxmlformats.org/presentationml/2006/main}modifyVerifier'):
                    root.remove(elem)
                tree.write(pres_xml)
            except Exception as e:
                print(f"Unlock error in {filename}: {e}")

        # Repack to pptx
        output_zip = os.path.join(temp_dir, "final.zip")
        shutil.make_archive(output_zip.replace('.zip', ''), 'zip', extract_path)

        # Rename .zip to .pptx
        final_pptx_path = output_zip.replace('.zip', '.pptx')
        os.rename(output_zip, final_pptx_path)

        with open(final_pptx_path, 'rb') as f:
            return f.read()
