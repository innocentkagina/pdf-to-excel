import os
import sys


sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from pdftoexcel import convert_pdfs
from transform import transform_excels
from consolidate import apply_final_transformations


def run_pdf_to_excel_converter():
    print("====================Starting converting PDFS====================")
    convert_pdfs()
    print("====================Done converting PDFS====================")

    print("====================Starting transforming Excel Files====================")
    transform_excels()
    print("====================Done transforming Excel Files====================")

    print("====================Starting Final transforming Excel Files====================")
    apply_final_transformations()
    print("====================Done Final transforming Excel Files====================")


if __name__ == "__main__":
    run_pdf_to_excel_converter()
