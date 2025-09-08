import os
import pandas as pd
from win32com import client

def convert_xlsx_to_pdf(input_file, output_file):
    # Open Excel application
    excel_app = client.Dispatch("Excel.Application")
    excel_app.Visible = False
    
    # Load the Excel file
    workbook = excel_app.Workbooks.Open(input_file)
    
    # Save as PDF
    workbook.ExportAsFixedFormat(0, output_file)
    
    # Close the workbook and quit Excel
    workbook.Close(False)
    excel_app.Quit()

def convert_docx_to_pdf(input_file, output_file):
    # Open Word application
    word_app = client.Dispatch("Word.Application")
    word_app.Visible = False
    
    # Load the Word document
    doc = word_app.Documents.Open(input_file)
    
    # Save as PDF
    doc.SaveAs(output_file, FileFormat=17)
    
    # Close the document and quit Word
    doc.Close(False)
    word_app.Quit()

def convert_to_pdf(input_file):
    # Determine the file extension
    file_extension = os.path.splitext(input_file)[1].lower()
    
    if file_extension == '.xlsx':
        output_file = input_file.replace('.xlsx', '.pdf')
        convert_xlsx_to_pdf(input_file, output_file)
        print(f"Converted {input_file} to {output_file}")
    
    elif file_extension == '.docx':
        output_file = input_file.replace('.docx', '.pdf')
        convert_docx_to_pdf(input_file, output_file)
        print(f"Converted {input_file} to {output_file}")
    
    else:
        print(f"Unsupported file format: {file_extension}")

# Example usage
convert_to_pdf("example.xlsx")
convert_to_pdf("example.docx")
