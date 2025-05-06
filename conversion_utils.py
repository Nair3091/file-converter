from pdf2docx import Converter
from PIL import Image
import os
from docx2pdf import convert as docx2pdf_convert

# Convert PDF to DOCX
def convert_pdf_to_docx(pdf_file_path):
    try:
        # Check if file exists
        if not os.path.exists(pdf_file_path):
            raise FileNotFoundError(f"The file {pdf_file_path} does not exist.")
        
        docx_file_path = pdf_file_path.replace('.pdf', '.docx')
        
        # Check if output file already exists to prevent overwriting
        if os.path.exists(docx_file_path):
            raise FileExistsError(f"The output file {docx_file_path} already exists.")
        
        # Perform conversion
        cv = Converter(pdf_file_path)
        cv.convert(docx_file_path, start=0, end=None)
        cv.close()
        
        return docx_file_path
    
    except Exception as e:
        raise Exception(f"Error converting PDF to DOCX: {str(e)}")

# Convert JPG to PNG
def convert_jpg_to_png(jpg_file_path):
    try:
        # Check if file exists
        if not os.path.exists(jpg_file_path):
            raise FileNotFoundError(f"The file {jpg_file_path} does not exist.")
        
        png_file_path = jpg_file_path.replace('.jpg', '.png')
        
        # Check if output file already exists to prevent overwriting
        if os.path.exists(png_file_path):
            raise FileExistsError(f"The output file {png_file_path} already exists.")
        
        # Perform conversion
        img = Image.open(jpg_file_path)
        img.save(png_file_path, 'PNG')
        
        return png_file_path
    
    except Exception as e:
        raise Exception(f"Error converting JPG to PNG: {str(e)}")

# Convert PNG to JPG
def convert_png_to_jpg(png_file_path):
    try:
        # Check if file exists
        if not os.path.exists(png_file_path):
            raise FileNotFoundError(f"The file {png_file_path} does not exist.")
        
        jpg_file_path = png_file_path.replace('.png', '.jpg')
        
        # Check if output file already exists to prevent overwriting
        if os.path.exists(jpg_file_path):
            raise FileExistsError(f"The output file {jpg_file_path} already exists.")
        
        # Perform conversion
        img = Image.open(png_file_path)
        img = img.convert('RGB')  # Convert PNG to RGB before saving as JPG
        img.save(jpg_file_path, 'JPEG')
        
        return jpg_file_path
    
    except Exception as e:
        raise Exception(f"Error converting PNG to JPG: {str(e)}")

# Convert DOCX to PDF
def convert_docx_to_pdf(docx_file_path):
    try:
        # Check if file exists
        if not os.path.exists(docx_file_path):
            raise FileNotFoundError(f"The file {docx_file_path} does not exist.")
        
        # Output PDF will be in the same directory
        output_folder = os.path.dirname(docx_file_path)
        
        # Check if output file already exists to prevent overwriting
        pdf_file_path = docx_file_path.replace('.docx', '.pdf')
        if os.path.exists(pdf_file_path):
            raise FileExistsError(f"The output file {pdf_file_path} already exists.")
        
        # Perform conversion
        docx2pdf_convert(docx_file_path, output_folder)
        
        return pdf_file_path
    
    except Exception as e:
        raise Exception(f"Error converting DOCX to PDF: {str(e)}")

