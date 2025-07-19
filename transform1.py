import os
import re
from pathlib import Path
import pandas as pd
from openpyxl.utils import get_column_letter
import logging

def extract_metadata(info_lines):
    """
    Extract district, parish, constituency, polling station, sub county info from list of header strings.
    """
    metadata = dict.fromkeys(['district', 'parish', 'constituency', 'polling_station', 'sub_county'], None)

    for line in info_lines:
        if 'District' in line and 'Parish' in line:
            # Example: District : 76 OYAM Parish : 01 ABANYA
            match = re.search(r'District\s*:\s*([^P]+)Parish\s*:\s*(.+)', line)
            if match:
                metadata['district'] = match.group(1).strip()
                metadata['parish'] = match.group(2).strip()
        elif 'Constituency' in line and 'Polling Station' in line:
            # Example: Constituency : 004 OYAM COUNTY NORTH Polling Station : 01 BAR OBIA
            match = re.search(r'Constituency\s*:\s*([^P]+)Polling Station\s*:\s*(.+)', line)
            if match:
                metadata['constituency'] = match.group(1).strip()
                metadata['polling_station'] = match.group(2).strip()
        elif 'Sub-County' in line:
            # Example: Sub-County : 01 ACHABA
            match = re.search(r'Sub-County\s*:\s*(.+)', line)
            if match:
                metadata['sub_county'] = match.group(1).strip()
    return metadata

def transform_excels_with_metadata(
    input_dir="converted",
    output_dir="transformed",
    skip_rows=3,
    serial_number_column="Serial No"
):
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s: %(message)s")

    for excel_file in Path(input_dir).glob("*.xlsx"):
        outname = Path(output_dir) / f"transformed_{excel_file.name}"
        processed_data = []
        logging.info(f"Processing file: {excel_file.name}")

        try:
            all_sheets = pd.read_excel(excel_file, engine="openpyxl", sheet_name=None)
        except Exception as e:
            logging.error(f"Could not read {excel_file.name}: {e}")
            continue

        for sheet, df in all_sheets.items():
            try:

                info_lines = df.iloc[3:5, 0].dropna().astype(str).tolist()
                metadata = extract_metadata(info_lines)

                # Skip first 3 rows (per latest requirement)
                df_clean = df.iloc[skip_rows:].reset_index(drop=True)

                # Set column headers (assume the new first row is the header)
                df_clean.columns = df_clean.iloc[0]
                df_clean = df_clean.iloc[1:].reset_index(drop=True)

                # Remove fully empty rows
                df_clean = df_clean.dropna(how='all')

                # Clean up whitespace from all string columns
                df_clean = df_clean.apply(
                    lambda col: col.str.strip() if col.dtype == "object" else col
                )

                # Add metadata columns
                for key, value in metadata.items():
                    df_clean[key] = value

                processed_data.append(df_clean)
            except Exception as e:
                logging.error(f"Error processing sheet {sheet}: {e}")

        # Concatenate all sheets data, add serial number, and save
        if processed_data:
            combined = pd.concat(processed_data, ignore_index=True)
            combined.insert(0, serial_number_column, range(1, len(combined) + 1))

            with pd.ExcelWriter(outname, engine="openpyxl") as writer:
                combined.to_excel(writer, index=False)
                worksheet = writer.sheets["Sheet1"]
                for idx, column in enumerate(combined.columns, 1):
                    max_width = max(combined[column].astype(str).map(len).max(), len(str(column))) + 2
                    worksheet.column_dimensions[get_column_letter(idx)].width = max_width

            logging.info(f"Exported cleaned data to: {outname}")


extract_metadata(info_lines)
# transform_excels_with_metadata(
#         input_dir="converted",    # Set to your source directory
#         output_dir="transformed", # Set to your output directory
#         skip_rows=5
#     )
