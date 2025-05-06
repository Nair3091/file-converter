from flask import Flask, request, send_file, render_template
import os
from werkzeug.utils import secure_filename
from PIL import Image
from docx import Document
from fpdf import FPDF  # alt for docx to pdf (if docx2pdf fails)
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from docx import Document

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return "No file part", 400

    file = request.files['file']
    if file.filename == '':
        return "No selected file", 400

    conversion_type = request.form.get('conversion_type')
    filename = secure_filename(file.filename)
    input_path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(input_path)

    # Determine output path
    base, ext = os.path.splitext(filename)
    output_path = os.path.join(UPLOAD_FOLDER, f"{base}_converted")

    try:
        if conversion_type == 'jpg-to-png' and ext.lower() in ['.jpg', '.jpeg']:
            output_path += '.png'
            Image.open(input_path).save(output_path)

        elif conversion_type == 'png-to-jpg' and ext.lower() == '.png':
            output_path += '.jpg'
            Image.open(input_path).convert("RGB").save(output_path)

        elif conversion_type == 'pdf-to-docx' and ext.lower() == '.pdf':
            # Placeholder: actual conversion requires pdf2docx or API
            return "PDF to DOCX conversion not yet implemented", 501

        elif conversion_type == 'docx-to-pdf' and ext.lower() == '.docx':
            output_path += '.pdf'
            convert_docx_to_pdf(input_path, output_path)

        else:
            return "Unsupported conversion or file type.", 400

        return send_file(output_path, as_attachment=True)

    except Exception as e:
        return f"Error during conversion: {str(e)}", 500

def convert_docx_to_pdf(docx_path, pdf_path):
# Open the DOCX document
    document = Document(docx_path)
    
    # Create a PDF canvas to write text
    c = canvas.Canvas(pdf_path, pagesize=letter)
    width, height = letter  # Standard letter size (8.5 x 11 inches)
    c.setFont("Helvetica", 12)

    # Starting position for the first paragraph
    y_position = height - 40  # Start at the top of the page
    
    # Add each paragraph from the DOCX
    for para in document.paragraphs:
        if y_position < 40:  # If there's no space left on the page, create a new page
            c.showPage()
            c.setFont("Helvetica", 12)
            y_position = height - 40
        
        # Write the paragraph text to the PDF
        c.drawString(40, y_position, para.text)
        y_position -= 14  # Move down for the next line
    
    # Save the generated PDF
    c.save()

if __name__ == "__main__":
    from waitress import serve
    serve(app, host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
