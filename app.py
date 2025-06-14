import streamlit as st
import os
from docx2pdf import convert
import tempfile
from pathlib import Path

# Set page config
st.set_page_config(
    page_title="Punjabi Text Converter",
    page_icon="📄",
    layout="centered"
)

# Custom CSS
st.markdown("""
    <style>
    .stButton>button {
        background: linear-gradient(90deg, #009688 60%, #4dd0e1 100%);
        color: white;
        border: none;
        padding: 0.5rem 1rem;
        border-radius: 0.5rem;
        font-weight: 600;
    }
    .stButton>button:hover {
        background: linear-gradient(90deg, #00796b 60%, #26c6da 100%);
    }
    .main {
        background: linear-gradient(120deg, #f5f5f5 60%, #e0f7fa 100%);
    }
    </style>
    """, unsafe_allow_html=True)

# Title
st.title("Punjabi Text Converter")
st.markdown("Convert your Word documents to PDF while preserving Punjabi text formatting")

# File uploader
uploaded_file = st.file_uploader("Choose a Word document", type=['docx'])

if uploaded_file is not None:
    # Create a temporary directory
    with tempfile.TemporaryDirectory() as temp_dir:
        # Save the uploaded file
        docx_path = os.path.join(temp_dir, uploaded_file.name)
        with open(docx_path, 'wb') as f:
            f.write(uploaded_file.getvalue())
        
        # Convert to PDF
        pdf_filename = Path(uploaded_file.name).stem + '.pdf'
        pdf_path = os.path.join(temp_dir, pdf_filename)
        
        try:
            with st.spinner('Converting... Please wait.'):
                convert(docx_path, pdf_path)
            
            # Read the PDF file
            with open(pdf_path, 'rb') as f:
                pdf_bytes = f.read()
            
            # Show success message
            st.success('Conversion complete!')
            
            # Display PDF preview
            st.markdown("### PDF Preview")
            st.components.v1.iframe(pdf_bytes, height=500)
            
            # Download button
            st.download_button(
                label="Download PDF",
                data=pdf_bytes,
                file_name=pdf_filename,
                mime="application/pdf"
            )
            
        except Exception as e:
            st.error(f'Error during conversion: {str(e)}') 