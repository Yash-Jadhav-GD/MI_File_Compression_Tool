import streamlit as st
from io import BytesIO
from PIL import Image
import pikepdf
import tempfile
import os
import comtypes.client
import zipfile

# ---------------- Compress images inside PDF ----------------
def compress_pdf_images(file_bytes, quality=30, max_width=1000, grayscale=False):
    file_bytes.seek(0)
    pdf = pikepdf.open(file_bytes)
    
    for page in pdf.pages:
        if not hasattr(page, "images"):
            continue
        for image_name, image_obj in page.images.items():
            try:
                raw_image = pdf.open_stream(image_obj)
                pil_img = Image.open(BytesIO(raw_image.read_bytes()))
                if pil_img.width > max_width:
                    ratio = max_width / pil_img.width
                    pil_img = pil_img.resize((max_width, int(pil_img.height * ratio)), Image.LANCZOS)
                if grayscale:
                    pil_img = pil_img.convert("L")
                img_bytes = BytesIO()
                pil_img.save(img_bytes, format="JPEG", quality=quality)
                img_bytes.seek(0)
                pdf.replace_image(image_obj, img_bytes.getvalue())
            except:
                continue
    out_bytes = BytesIO()
    pdf.save(out_bytes)
    out_bytes.seek(0)
    return out_bytes

# ---------------- Convert PPT to PDF ----------------
def ppt_to_pdf(file_bytes):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp_ppt:
        tmp_ppt.write(file_bytes.read())
        tmp_ppt_path = tmp_ppt.name
    tmp_pdf_path = tmp_ppt_path.replace(".pptx", ".pdf")
    try:
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        powerpoint.Visible = 1
        presentation = powerpoint.Presentations.Open(tmp_ppt_path, WithWindow=False)
        presentation.ExportAsFixedFormat(tmp_pdf_path, 2)
        presentation.Close()
        powerpoint.Quit()
        with open(tmp_pdf_path, "rb") as f:
            pdf_bytes = BytesIO(f.read())
        return pdf_bytes
    finally:
        if os.path.exists(tmp_ppt_path):
            os.remove(tmp_ppt_path)
        if os.path.exists(tmp_pdf_path):
            os.remove(tmp_pdf_path)

# ---------------- Streamlit UI ----------------
st.title("PDF & PPT Compression Tool")

file_type = st.radio("Select File Type:", ["PDF", "PPT"])
aggressiveness = st.slider("Compression Aggressiveness (lower = more compression)", 10, 90, 50)
grayscale_option = st.checkbox("Convert images to grayscale", value=False)

uploaded_files = st.file_uploader(
    "Upload files", accept_multiple_files=True,
    type=['pdf'] if file_type == "PDF" else ['ppt', 'pptx']
)

if uploaded_files:
    compressed_files = []
    for file in uploaded_files:
        original_size = round(len(file.getvalue())/1024, 2)
        try:
            if file_type == "PDF":
                compressed_file = compress_pdf_images(file, quality=aggressiveness, grayscale=grayscale_option)
                filename = f"compressed_{file.name}"
            else:
                pdf_bytes = ppt_to_pdf(file)
                compressed_file = pdf_bytes
                filename = f"{file.name.rsplit('.',1)[0]}.pdf"

            compressed_size = round(len(compressed_file.getvalue())/1024, 2)
            reduction = round(100*(original_size - compressed_size)/original_size, 2)
            st.success(f"{file.name}: Original {original_size} KB → Compressed {compressed_size} KB | Reduction: {reduction}%")

            compressed_files.append((filename, compressed_file.getvalue()))
        except Exception as e:
            st.error(f"Failed to process {file.name}: {e}")

    if compressed_files:
        # Create in-memory ZIP
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for fname, data in compressed_files:
                zip_file.writestr(fname, data)
        zip_buffer.seek(0)
        st.download_button(
            label="Download All Compressed Files (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="compressed_files.zip"
        )
