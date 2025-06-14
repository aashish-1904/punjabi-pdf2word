from flask import Flask, request, render_template, send_file, flash, redirect, url_for
import os
from werkzeug.utils import secure_filename
import tempfile
from docx2pdf import convert as docx2pdf_convert

app = Flask(__name__)
app.secret_key = 'supersecretkey'  # Needed for flashing messages
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()

ALLOWED_EXTENSIONS = {'docx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return 'No file uploaded', 400
    file = request.files['file']
    if file.filename == '':
        return 'No file selected', 400
    if file and allowed_file(file.filename):
        docx_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename))
        file.save(docx_path)
        pdf_filename = secure_filename(file.filename.rsplit('.', 1)[0] + '.pdf')
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf_filename)
        try:
            docx2pdf_convert(docx_path, pdf_path)
        except Exception as e:
            if os.path.exists(docx_path):
                os.remove(docx_path)
            return f'Conversion failed: {str(e)}', 500
        if os.path.exists(docx_path):
            os.remove(docx_path)
        if not os.path.exists(pdf_path):
            return 'PDF not created', 500
        return send_file(pdf_path, as_attachment=True, download_name=pdf_filename)
    return 'Invalid file type', 400

if __name__ == '__main__':
    app.run(debug=True) 