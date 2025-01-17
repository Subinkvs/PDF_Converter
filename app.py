import os
from flask import Flask, request, send_file
import pandas as pd
from docx import Document
from pptx import Presentation
from openpyxl import load_workbook
import subprocess
from tempfile import NamedTemporaryFile

app = Flask(__name__)

# Convert .pptx, .docx, .xlsx, .csv to PDF
def convert_to_pdf(input_file, file_type):
    # Temporary file to store the output PDF
    output_pdf = NamedTemporaryFile(delete=False, suffix=".pdf")
    output_pdf.close()
    
    if file_type == "pptx":
        # Convert PowerPoint (.pptx) to PDF using unoconv
        subprocess.run(["unoconv", "-f", "pdf", "-o", output_pdf.name, input_file])
    elif file_type == "docx":
        # Convert Word (.docx) to PDF using unoconv
        subprocess.run(["unoconv", "-f", "pdf", "-o", output_pdf.name, input_file])
    elif file_type in ["xlsx", "csv"]:
        # Convert Excel (.xlsx) or CSV to PDF using unoconv
        if file_type == "csv":
            # Convert CSV to Excel first
            temp_excel = NamedTemporaryFile(delete=False, suffix=".xlsx")
            df = pd.read_csv(input_file)
            df.to_excel(temp_excel.name, index=False)
            temp_excel.close()
            subprocess.run(["unoconv", "-f", "pdf", "-o", output_pdf.name, temp_excel.name])
            os.remove(temp_excel.name)
        else:
            subprocess.run(["unoconv", "-f", "pdf", "-o", output_pdf.name, input_file])
    
    return output_pdf.name

@app.route('/convert', methods=['POST'])
def convert_file():
    # Check if file is in the request
    if 'file' not in request.files:
        return "No file part", 400
    
    file = request.files['file']
    if file.filename == '':
        return "No selected file", 400
    
    # Save the uploaded file temporarily
    temp_file = NamedTemporaryFile(delete=False)
    file.save(temp_file.name)
    
    # Identify file type based on extension
    file_type = None
    if file.filename.endswith('.pptx'):
        file_type = 'pptx'
    elif file.filename.endswith('.docx'):
        file_type = 'docx'
    elif file.filename.endswith('.xlsx'):
        file_type = 'xlsx'
    elif file.filename.endswith('.csv'):
        file_type = 'csv'
    else:
        return "Unsupported file type", 400
    
    # Convert the file to PDF
    pdf_file = convert_to_pdf(temp_file.name, file_type)
    os.remove(temp_file.name)
    
    # Return the PDF file to the browser
    return send_file(
        pdf_file,
        as_attachment=True,
        download_name="converted.pdf",
        mimetype="application/pdf"
    )

if __name__ == "__main__":
    app.run(debug=True)
