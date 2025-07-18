from merge import merge_excels
from consolidate import apply_final_transformations
from transform import transform_excels
from pdftoexcel import convert_pdfs
from rename import rename_merged_files
from export_to_db import export_excels_to_postgres
import os
import sys


sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Example mapping: update this with your real info!
identifier_to_district = {
    "76": "Nairobi",
    "88": "Mombasa",
    "23": "Kisumu",
    # ...etc
}

RENAME_SUFFIX = '.xlsx'
# Set to True to overwrite if exists
OVERWRITE = False
FINAL_DIR = "merged"


def run_pdf_to_excel_converter():
    print("====================Starting converting PDFS====================")
    # convert_pdfs()
    print("====================Done converting PDFS====================")

    print("====================Starting transforming Excel Files====================")
    # transform_excels()
    print("====================Done transforming Excel Files====================")

    print("====================Starting Final transforming Excel Files====================")
    # apply_final_transformations()
    print("====================Done Final transforming Excel Files====================")

    print("====================Starting Merging Excel Files====================")
    # merge_excels()
    print("====================Done Merging Excel Files====================")

    print("====================Starting Renaming Excel Files====================")
    rename_merged_files(identifier_to_district,
                        ext=RENAME_SUFFIX, overwrite=OVERWRITE)
    print("====================Done Renaming Excel Files====================")

    print("====================Starting Exporting Excel Files To DB ====================")
    export_excels_to_postgres(FINAL_DIR)
    print("====================Done Exporting Excel Files To DB ====================")


if __name__ == "__main__":
    run_pdf_to_excel_converter()
