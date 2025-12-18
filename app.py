import streamlit as st
import pandas as pd
from docx import Document
import io
import zipfile
import os

st.set_page_config(page_title="Doc Generator", layout="centered")

st.title("📄 Personalized Document Generator")
st.write("Upload your Excel/CSV and Word Template to generate individual files.")

# --- File Uploaders ---
col1, col2 = st.columns(2)
with col1:
    csv_file = st.file_uploader("Upload CSV/Excel", type=['csv', 'xlsx'])
with col2:
    template_file = st.file_uploader("Upload Word Template", type=['docx'])

if csv_file and template_file:
    # Read Data
    if csv_file.name.endswith('.csv'):
        df = pd.read_csv(csv_file)
    else:
        df = pd.read_excel(csv_file)

    # Clean data
    df = df.dropna(subset=['Student Name'])
    
    st.success(f"Loaded {len(df)} students.")

    if st.button("🚀 Generate Documents"):
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for index, row in df.iterrows():
                # Load template from memory
                template_file.seek(0)
                doc = Document(template_file)
                
                replacements = {
                    "<Student Name>": str(row['Student Name']),
                    "<College Name>": str(row['College Name'])
                }
                
                # Replacement Logic
                for p in doc.paragraphs:
                    for placeholder, value in replacements.items():
                        if placeholder in p.text:
                            p.text = p.text.replace(placeholder, value)
                
                for table in doc.tables:
                    for row_tab in table.rows:
                        for cell in row_tab.cells:
                            for p in cell.paragraphs:
                                for placeholder, value in replacements.items():
                                    if placeholder in p.text:
                                        p.text = p.text.replace(placeholder, value)

                # Save doc to memory buffer
                doc_io = io.BytesIO()
                doc.save(doc_io)
                doc_io.seek(0)
                
                # Add to ZIP
                file_name = f"{str(row['Student Name']).replace(' ', '_')}.docx"
                zip_file.writestr(file_name, doc_io.getvalue())

        st.balloons()
        st.download_button(
            label="📥 Download All (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="Personalized_Documents.zip",
            mime="application/zip"
        )