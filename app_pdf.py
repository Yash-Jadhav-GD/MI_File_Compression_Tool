import fitz  # PyMuPDF
import streamlit as st
from io import BytesIO
import zipfile

def compress_pdf_strong(file_bytes, quality=40):
    doc = fitz.open(stream=file_bytes.read(), filetype="pdf")
    out_bytes = BytesIO()

    # Reduce image quality
    for page in doc:
        for img in page.get_images(full=True):
            xref = img[0]
            pix = fitz.Pixmap(doc, xref)
            if pix.n > 4:  # CMYK -> convert to RGB
                pix = fitz.Pixmap(fitz.csRGB, pix)
            pix.save(f"temp.jpg", quality=quality)  # recompress as JPEG
            doc.update_image(xref, open("temp.jpg", "rb").read())

    doc.save(out_bytes, deflate=True)
    out_bytes.seek(0)
    return out_bytes

# Streamlit UI
st.title("Market Intelligence PDF Compression Tool")

uploaded_files = st.file_uploader("Upload PDF files", accept_multiple_files=True, type=['pdf'])

if uploaded_files:
    compressed_files = []
    for file in uploaded_files:
        original_size = round(len(file.getvalue())/1024, 2)
        file.seek(0)
        compressed_file = compress_pdf_strong(file, quality=40)

        filename = f"compressed_{file.name}"
        compressed_size = round(len(compressed_file.getvalue())/1024, 2)
        reduction = round(100*(original_size - compressed_size)/original_size, 2)

        st.success(f"{file.name}: Original {original_size} KB → Compressed {compressed_size} KB | Reduction: {reduction}%")
        compressed_files.append((filename, compressed_file.getvalue()))

    if compressed_files:
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for fname, data in compressed_files:
                zip_file.writestr(fname, data)
        zip_buffer.seek(0)

        st.download_button(
            label="Download All Compressed PDFs (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="compressed_pdfs.zip"
        )

