from flask import Flask, request, send_file, render_template
import os
from conversion_utils import convert_pdf_to_docx, convert_jpg_to_png, convert_png_to_jpg, convert_docx_to_pdf

app = Flask(__name__)

# Set upload and static directories
UPLOAD_FOLDER = 'uploads'
STATIC_FOLDER = 'static'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['STATIC_FOLDER'] = STATIC_FOLDER

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert_file():
    file = request.files['file']
    conversion_type = request.form['conversion_type']

    if not file:
        return "No file uploaded", 400

    file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(file_path)

    converted_file_path = ""

    try:
        if conversion_type == 'pdf-to-docx':
            converted_file_path = convert_pdf_to_docx(file_path)
        elif conversion_type == 'jpg-to-png':
            converted_file_path = convert_jpg_to_png(file_path)
        elif conversion_type == 'png-to-jpg':
            converted_file_path = convert_png_to_jpg(file_path)
        elif conversion_type == 'docx-to-pdf':
            converted_file_path = convert_docx_to_pdf(file_path)
        else:
            return "Unsupported conversion type", 400
        
        return send_file(converted_file_path, as_attachment=True)
    
    finally:
        os.remove(file_path)  # Clean up the uploaded file

if __name__ == "__main__":
    from waitress import serve
    serve(app, host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
