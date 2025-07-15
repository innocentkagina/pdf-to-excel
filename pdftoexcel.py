import tabula
import pandas as pd
import os
from pathlib import Path

def convert_pdfs(input_dir="original", output_dir="converted"):
    # Create output directory if it doesn't exist
    Path(output_dir).mkdir(parents=True, exist_ok=True)

    # Get list of PDF files in input directory
    pdf_files = [f for f in os.listdir(input_dir) if f.lower().endswith('.pdf')]

    for pdf_file in pdf_files:
        input_path = os.path.join(input_dir, pdf_file)
        output_path = os.path.join(output_dir, f"{Path(pdf_file).stem}.xlsx")

        # Convert PDF to Excel
        pdf_to_excel(input_path, output_path)
        print(f"Converted: {pdf_file} -> {output_path}")

def pdf_to_excel(pdf_file_path, excel_file_path):
    # Read PDF file
    tables = tabula.read_pdf(pdf_file_path, pages='all')

    # Write each table to a separate sheet in the Excel file
    with pd.ExcelWriter(excel_file_path) as writer:
        for i, table in enumerate(tables):
            table.to_excel(writer, sheet_name=f'Sheet{i+1}')
