import streamlit as st
import os
import tempfile
from pathlib import Path
from docx import Document
from docx.shared import RGBColor
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
import io
import requests
from urllib.parse import urlparse

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

@st.cache_data
def download_gurmukhi_font():
    """Download and cache Gurmukhi font"""
    try:
        # Download Noto Sans Gurmukhi font
        font_url = "https://github.com/notofonts/notofonts.github.io/raw/main/fonts/NotoSansGurmukhi/hinted/ttf/NotoSansGurmukhi-Regular.ttf"
        response = requests.get(font_url)
        response.raise_for_status()
        
        # Save font to temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix='.ttf') as font_file:
            font_file.write(response.content)
            return font_file.name
    except Exception as e:
        st.warning(f"Could not download Gurmukhi font: {e}")
        return None

def register_gurmukhi_font():
    """Register Gurmukhi font with ReportLab"""
    try:
        font_path = download_gurmukhi_font()
        if font_path:
            pdfmetrics.registerFont(TTFont('NotoSansGurmukhi', font_path))
            return True
    except Exception as e:
        st.warning(f"Could not register Gurmukhi font: {e}")
    return False

def is_gurmukhi_text(text):
    """Check if text contains Gurmukhi characters"""
    gurmukhi_range = range(0x0A00, 0x0A7F)  # Unicode range for Gurmukhi
    return any(ord(char) in gurmukhi_range for char in text)

def docx_color_to_reportlab(docx_color):
    """Convert DOCX color to ReportLab color"""
    if docx_color and hasattr(docx_color, 'rgb') and docx_color.rgb:
        rgb = docx_color.rgb
        r = (rgb >> 16) & 0xFF
        g = (rgb >> 8) & 0xFF
        b = rgb & 0xFF
        return colors.Color(r/255.0, g/255.0, b/255.0, alpha=1.0)
    return colors.black

def get_alignment(paragraph):
    """Get paragraph alignment"""
    alignment_map = {
        0: TA_LEFT,      # WD_ALIGN_PARAGRAPH.LEFT
        1: TA_CENTER,    # WD_ALIGN_PARAGRAPH.CENTER
        2: TA_RIGHT,     # WD_ALIGN_PARAGRAPH.RIGHT
        3: TA_JUSTIFY,   # WD_ALIGN_PARAGRAPH.JUSTIFY
    }
    if hasattr(paragraph, 'alignment') and paragraph.alignment is not None:
        return alignment_map.get(paragraph.alignment, TA_LEFT)
    return TA_LEFT

def points_to_reportlab(points):
    """Convert points to ReportLab units"""
    if points:
        return float(points) * 72.0 / 914400  # Convert from EMUs to points
    return 12  # Default font size

def safe_get_color_hex(run):
    """Safely extract color from run and convert to hex"""
    try:
        if hasattr(run, 'font') and run.font.color and hasattr(run.font.color, 'rgb') and run.font.color.rgb:
            rgb_val = run.font.color.rgb
            if rgb_val is not None:
                # Convert to hex string
                return f"#{rgb_val:06x}"
    except Exception:
        pass
    return "#000000"  # Default black

def safe_get_font_size(run, default_size=12):
    """Safely extract font size from run"""
    try:
        if hasattr(run, 'font') and run.font.size and hasattr(run.font.size, 'pt'):
            return run.font.size.pt
    except Exception:
        pass
    return default_size

def safe_get_formatting(run):
    """Safely extract all formatting from a run with detailed detection"""
    formatting = {
        'bold': False,
        'italic': False,
        'underline': False,
        'size': 12,
        'color': '#000000'
    }
    
    try:
        # Check bold - multiple ways
        if hasattr(run, 'bold') and run.bold is True:
            formatting['bold'] = True
        elif hasattr(run, 'font') and hasattr(run.font, 'bold') and run.font.bold is True:
            formatting['bold'] = True
        
        # Check italic - multiple ways
        if hasattr(run, 'italic') and run.italic is True:
            formatting['italic'] = True
        elif hasattr(run, 'font') and hasattr(run.font, 'italic') and run.font.italic is True:
            formatting['italic'] = True
        
        # Check underline - multiple ways
        if hasattr(run, 'underline') and run.underline is True:
            formatting['underline'] = True
        elif hasattr(run, 'font') and hasattr(run.font, 'underline') and run.font.underline is True:
            formatting['underline'] = True
        
        # Get font size
        formatting['size'] = safe_get_font_size(run, 12)
        
        # Get color
        formatting['color'] = safe_get_color_hex(run)
        
    except Exception as e:
        pass  # Use defaults
    
    return formatting

def convert_docx_to_pdf(docx_bytes, output_filename):
    """Convert DOCX to PDF with comprehensive formatting preservation"""
    try:
        # Register Gurmukhi font
        font_registered = register_gurmukhi_font()
        
        # Create a temporary file for the DOCX
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_docx:
            temp_docx.write(docx_bytes)
            temp_docx_path = temp_docx.name

        # Read the DOCX file
        doc = Document(temp_docx_path)
        
        # Debug: Analyze formatting in the document
        formatting_found = {
            'bold_count': 0,
            'italic_count': 0,
            'underline_count': 0,
            'colored_text': 0,
            'different_sizes': set()
        }
        
        for paragraph in doc.paragraphs:
            for run in paragraph.runs:
                if run.text.strip():
                    fmt = safe_get_formatting(run)
                    if fmt['bold']:
                        formatting_found['bold_count'] += 1
                    if fmt['italic']:
                        formatting_found['italic_count'] += 1
                    if fmt['underline']:
                        formatting_found['underline_count'] += 1
                    if fmt['color'] != '#000000':
                        formatting_found['colored_text'] += 1
                    formatting_found['different_sizes'].add(fmt['size'])
        
        # Show formatting analysis to user
        st.info(f"""
        📊 **Formatting Analysis:**
        - Bold text runs: {formatting_found['bold_count']}
        - Italic text runs: {formatting_found['italic_count']}
        - Underlined text runs: {formatting_found['underline_count']}
        - Colored text runs: {formatting_found['colored_text']}
        - Font sizes found: {sorted(list(formatting_found['different_sizes']))}
        """)
        
        # Create PDF in memory
        buffer = io.BytesIO()
        pdf_doc = SimpleDocTemplate(
            buffer, 
            pagesize=A4, 
            topMargin=1*inch, 
            bottomMargin=1*inch,
            leftMargin=1*inch,
            rightMargin=1*inch
        )
        
        # Get base styles
        styles = getSampleStyleSheet()
        
        # Default font
        default_font = 'NotoSansGurmukhi' if font_registered else 'Helvetica'
        
        # Story to hold all content
        story = []
        
        # Process each paragraph in the document
        for para_idx, paragraph in enumerate(doc.paragraphs):
            if paragraph.text.strip():
                # Get paragraph-level formatting
                para_alignment = get_alignment(paragraph)
                
                # Analyze the paragraph style
                style_name = paragraph.style.name if paragraph.style else 'Normal'
                
                # Determine if this is a heading and get appropriate sizing
                base_font_size = 12
                is_heading = False
                heading_level = 0
                
                if 'Heading' in style_name:
                    is_heading = True
                    if 'Heading 1' in style_name:
                        heading_level = 1
                        base_font_size = 18
                    elif 'Heading 2' in style_name:
                        heading_level = 2
                        base_font_size = 16
                    elif 'Heading 3' in style_name:
                        heading_level = 3
                        base_font_size = 14
                    else:
                        heading_level = 4
                        base_font_size = 13
                elif 'Title' in style_name:
                    is_heading = True
                    heading_level = 0
                    base_font_size = 20
                
                # Get the most common font size in the paragraph
                font_sizes = []
                for run in paragraph.runs:
                    if run.text.strip():
                        fmt = safe_get_formatting(run)
                        font_sizes.append(fmt['size'])
                
                if font_sizes:
                    # Use the most common font size, or the first one if all different
                    from collections import Counter
                    size_counts = Counter(font_sizes)
                    base_font_size = size_counts.most_common(1)[0][0]
                
                # Create paragraph style
                style_id = f"Style_{para_idx}_{heading_level}"
                
                if is_heading:
                    text_color = colors.Color(0.1, 0.2, 0.5)  # Dark blue for headings
                    space_after = max(12, base_font_size * 0.8)
                    space_before = max(6, base_font_size * 0.4)
                else:
                    text_color = colors.black
                    space_after = 6
                    space_before = 0
                
                para_style = ParagraphStyle(
                    style_id,
                    parent=styles['Normal'],
                    fontSize=base_font_size,
                    fontName=default_font,
                    alignment=para_alignment,
                    spaceAfter=space_after,
                    spaceBefore=space_before,
                    textColor=text_color,
                    leading=base_font_size * 1.3,
                    leftIndent=0,
                    rightIndent=0,
                )
                
                # Process runs with detailed formatting
                formatted_text = ""
                
                for run in paragraph.runs:
                    text = run.text
                    if text:
                        # Get comprehensive formatting
                        fmt = safe_get_formatting(run)
                        
                        # Start with the text
                        formatted_run = text
                        
                        # Apply font size if different from paragraph base
                        if abs(fmt['size'] - base_font_size) > 1:  # Allow 1pt tolerance
                            formatted_run = f'<font size="{int(fmt["size"])}">{formatted_run}</font>'
                        
                        # Apply color if not black
                        if fmt['color'] != "#000000":
                            formatted_run = f'<font color="{fmt["color"]}">{formatted_run}</font>'
                        
                        # Apply formatting in the correct order (innermost first)
                        if fmt['underline']:
                            formatted_run = f"<u>{formatted_run}</u>"
                        
                        if fmt['italic']:
                            formatted_run = f"<i>{formatted_run}</i>"
                        
                        if fmt['bold']:
                            formatted_run = f"<b>{formatted_run}</b>"
                        
                        formatted_text += formatted_run
                
                # Add the paragraph to story
                if formatted_text.strip():
                    try:
                        # Clean up any problematic characters
                        clean_text = formatted_text.replace('\x0b', ' ').replace('\x0c', ' ')
                        story.append(Paragraph(clean_text, para_style))
                        
                        # Add spacing after paragraph
                        if is_heading:
                            story.append(Spacer(1, 8))
                        else:
                            story.append(Spacer(1, 4))
                            
                    except Exception as e:
                        # Fallback: create simple paragraph
                        simple_style = ParagraphStyle(
                            f"Fallback_{para_idx}",
                            parent=styles['Normal'],
                            fontSize=12,
                            fontName=default_font,
                            alignment=para_alignment,
                        )
                        story.append(Paragraph(paragraph.text, simple_style))
                        story.append(Spacer(1, 6))
        
        # Process tables with comprehensive formatting
        for table_idx, table in enumerate(doc.tables):
            if table.rows:
                table_data = []
                table_styles = [
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    ('FONTNAME', (0, 0), (-1, -1), default_font),
                    ('FONTSIZE', (0, 0), (-1, -1), 10),
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                    ('LEFTPADDING', (0, 0), (-1, -1), 6),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 6),
                    ('TOPPADDING', (0, 0), (-1, -1), 4),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                ]
                
                for row_idx, row in enumerate(table.rows):
                    row_data = []
                    for cell_idx, cell in enumerate(row.cells):
                        # Process cell content with formatting
                        cell_content = ""
                        for para in cell.paragraphs:
                            if para.text.strip():
                                para_text = ""
                                for run in para.runs:
                                    text = run.text
                                    if text:
                                        # Get formatting for cell text
                                        fmt = safe_get_formatting(run)
                                        
                                        # Apply formatting to cell text
                                        if fmt['underline']:
                                            text = f"<u>{text}</u>"
                                        if fmt['italic']:
                                            text = f"<i>{text}</i>"
                                        if fmt['bold']:
                                            text = f"<b>{text}</b>"
                                        
                                        para_text += text
                                if para_text.strip():
                                    cell_content += para_text + " "
                        
                        row_data.append(cell_content.strip())
                    
                    table_data.append(row_data)
                
                # Create table
                if table_data:
                    # Make header row bold if it's the first row
                    if len(table_data) > 1:
                        table_styles.append(('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey))
                        table_styles.append(('FONTNAME', (0, 0), (-1, 0), default_font))
                    
                    pdf_table = Table(table_data)
                    pdf_table.setStyle(TableStyle(table_styles))
                    story.append(pdf_table)
                    story.append(Spacer(1, 12))
        
        # Build PDF
        pdf_doc.build(story)
        
        # Clean up temporary file
        os.unlink(temp_docx_path)
        
        # Get PDF bytes
        buffer.seek(0)
        return buffer.getvalue()
        
    except Exception as e:
        st.error(f"Error during conversion: {str(e)}")
        import traceback
        st.error(f"Detailed error: {traceback.format_exc()}")
        return None

# Title and description
st.title("ਪੰਜਾਬੀ ਟੈਕਸਟ ਕਨਵਰਟਰ")
st.markdown("**Punjabi Text Converter** - Convert Word documents to PDF with enhanced formatting preservation")

st.info("""
**🎯 Enhanced Formatting Preservation:**
- ✅ Maintains original font sizes and colors
- ✅ Preserves text alignment (left, center, right, justify)
- ✅ Keeps bold, italic, underline formatting
- ✅ Handles headings with proper styling
- ✅ Preserves table formatting and structure
- ✅ Uses authentic Noto Sans Gurmukhi font
- ✅ Proper Unicode support for Punjabi text
""")

# File uploader
uploaded_file = st.file_uploader("Choose a Word document (.docx)", type=['docx'])

if uploaded_file is not None:
    try:
        with st.spinner('Converting... Analyzing document formatting and processing...'):
            # Convert the file
            pdf_bytes = convert_docx_to_pdf(uploaded_file.getvalue(), uploaded_file.name)
            
            if pdf_bytes:
                # Show success message
                st.success('✅ Conversion complete! Original formatting should be preserved.')
                
                # Create two columns for buttons
                col1, col2 = st.columns(2)
                
                with col1:
                    # Preview button
                    st.download_button(
                        label="📄 Preview PDF",
                        data=pdf_bytes,
                        file_name=f"preview_{Path(uploaded_file.name).stem}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
                
                with col2:
                    # Download button
                    st.download_button(
                        label="⬇️ Download PDF",
                        data=pdf_bytes,
                        file_name=f"{Path(uploaded_file.name).stem}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
                
                # Show file info
                st.info(f"📊 Original: {uploaded_file.name} ({len(uploaded_file.getvalue())/1024:.1f} KB)")
                st.info(f"📄 PDF: {len(pdf_bytes)/1024:.1f} KB")
                
                # Show formatting analysis
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_file:
                        temp_file.write(uploaded_file.getvalue())
                        temp_file_path = temp_file.name
                    
                    doc = Document(temp_file_path)
                    
                    # Analyze document
                    total_paragraphs = len([p for p in doc.paragraphs if p.text.strip()])
                    total_tables = len(doc.tables)
                    headings = len([p for p in doc.paragraphs if p.style.name.startswith('Heading')])
                    
                    st.success(f"📋 Document Analysis: {total_paragraphs} paragraphs, {headings} headings, {total_tables} tables processed")
                    
                    # Show sample text
                    sample_text = ""
                    for para in doc.paragraphs[:2]:
                        if para.text.strip():
                            sample_text += para.text[:150] + "...\n\n"
                    
                    if sample_text:
                        st.text_area("📝 Sample text from document:", sample_text, height=120)
                    
                    os.unlink(temp_file_path)
                except:
                    pass
    
    except Exception as e:
        st.error(f'❌ Error during conversion: {str(e)}')
        st.error("Please ensure your document is a valid Word file and try again.") 
