from pdf2docx import Converter
from PIL import Image
import os

# Convert PDF to DOCX
def convert_pdf_to_docx(pdf_file_path):
    docx_file_path = pdf_file_path.replace('.pdf', '.docx')
    cv = Converter(pdf_file_path)
    cv.convert(docx_file_path, start=0, end=None)
    cv.close()
    return docx_file_path

# Convert JPG to PNG
def convert_jpg_to_png(jpg_file_path):
    png_file_path = jpg_file_path.replace('.jpg', '.png')
    img = Image.open(jpg_file_path)
    img.save(png_file_path, 'PNG')
    return png_file_path

# Convert PNG to JPG
def convert_png_to_jpg(png_file_path):
    jpg_file_path = png_file_path.replace('.png', '.jpg')
    img = Image.open(png_file_path)
    img = img.convert('RGB')  # Convert PNG to RGB before saving as JPG
    img.save(jpg_file_path, 'JPEG')
    return jpg_file_path

from docx2pdf import convert as docx2pdf_convert

def convert_docx_to_pdf(docx_file_path):
    # Output PDF will be in the same directory
    output_folder = os.path.dirname(docx_file_path)
    docx2pdf_convert(docx_file_path, output_folder)
    pdf_file_path = docx_file_path.replace('.docx', '.pdf')
    return pdf_file_path
