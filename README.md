# Punjabi Text Converter

A web application that converts Punjabi text in Gurmukhi from Word documents (.docx) to PDF format while preserving the text formatting.

## Features

- Upload Word documents containing Punjabi text
- Convert to PDF while maintaining text formatting
- Simple and intuitive web interface
- Secure file handling

## Prerequisites

- Python 3.7 or higher
- pip (Python package installer)

## Installation

1. Clone this repository:
```bash
git clone <repository-url>
cd translator
```

2. Create a virtual environment (recommended):
```bash
python -m venv venv
source venv/bin/activate  # On Windows, use: venv\Scripts\activate
```

3. Install the required packages:
```bash
pip install -r requirements.txt
```

4. Download a Gurmukhi font (e.g., from Google Fonts) and place it in the project directory.

5. Update the font path in `app.py`:
```python
pdfmetrics.registerFont(TTFont('Gurmukhi', 'path_to_your_gurmukhi_font.ttf'))
```

## Running the Application

1. Start the Flask development server:
```bash
python app.py
```

2. Open your web browser and navigate to:
```
http://localhost:5000
```

## Usage

1. Click on the upload area or drag and drop a Word document (.docx) containing Punjabi text
2. Click the "Convert to PDF" button
3. Wait for the conversion to complete
4. The PDF file will be automatically downloaded

## Notes

- Maximum file size: 16MB
- Only .docx files are supported
- The application preserves the text formatting during conversion
- Temporary files are automatically cleaned up after conversion

## License

This project is licensed under the MIT License - see the LICENSE file for details. 