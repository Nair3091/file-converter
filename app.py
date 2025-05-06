from flask import Flask, request, send_file, render_template
import os
from werkzeug.utils import secure_filename
from PIL import Image
from docx import Document
from fpdf import FPDF  # alt for docx to pdf (if docx2pdf fails)
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from pdf2docx import Converter  # for PDF to DOCX conversion

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
        # Handle image conversions
        if conversion_type == 'jpg-to-png' and ext.lower() in ['.jpg', '.jpeg']:
            output_path += '.png'
            Image.open(input_path).save(output_path)

        elif conversion_type == 'png-to-jpg' and ext.lower() == '.png':
            output_path += '.jpg'
            Image.open(input_path).convert("RGB").save(output_path)

        # Handle PDF to DOCX conversion
        elif conversion_type == 'pdf-to-docx' and ext.lower() == '.pdf':
            output_path += '.docx'
            convert_pdf_to_docx(input_path, output_path)

        # Handle DOCX to PDF conversion
        elif conversion_type == 'docx-to-pdf' and ext.lower() == '.docx':
            output_path += '.pdf'
            convert_docx_to_pdf(input_path, output_path)

        else:
            return "Unsupported conversion or file type.", 400

        # Return the converted file
        return send_file(output_path, as_attachment=True)

    except Exception as e:
        return f"Error during conversion: {str(e)}", 500


def convert_pdf_to_docx(pdf_path, docx_path):
    """
    Converts a PDF file to DOCX using pdf2docx.
    """
    try:
        cv = Converter(pdf_path)
        cv.convert(docx_path)
        cv.close()
    except Exception as e:
        raise Exception(f"Error converting PDF to DOCX: {str(e)}")


def convert_docx_to_pdf(docx_path, pdf_path):
    """
    Converts a DOCX file to PDF.
    """
    try:
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
    except Exception as e:
        raise Exception(f"Error converting DOCX to PDF: {str(e)}")


if __name__ == "__main__":
    from waitress import serve
    serve(app, host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
