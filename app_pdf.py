import streamlit as st
from io import BytesIO
from PyPDF2 import PdfReader, PdfWriter
import zipfile

def compress_pdf(file_bytes):
    file_bytes.seek(0)
    reader = PdfReader(file_bytes)
    writer = PdfWriter()

    for page in reader.pages:
        writer.add_page(page)

    # write compressed PDF
    out_bytes = BytesIO()
    writer.write(out_bytes)
    out_bytes.seek(0)
    return out_bytes

# Streamlit UI
st.title("PDF Compression Tool")

uploaded_files = st.file_uploader("Upload PDF files", accept_multiple_files=True, type=['pdf'])

if uploaded_files:
    compressed_files = []

    for file in uploaded_files:
        original_size = round(len(file.getvalue()) / 1024, 2)
        compressed_file = compress_pdf(file)

        filename = f"compressed_{file.name}"
        compressed_size = round(len(compressed_file.getvalue()) / 1024, 2)
        reduction = round(100 * (original_size - compressed_size) / original_size, 2)

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
