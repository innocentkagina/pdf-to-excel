import os
import re
import logging
import pandas as pd
from typing import List, Optional, Dict

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s %(levelname)s: %(message)s")

ROW_LIMIT_PER_SHEET = 1_024_000


def extract_identifier(filename: str) -> Optional[str]:
    match = re.search(r'NVR_REGISTER_TXT_([^_\.]+)', filename)
    return match.group(1) if match else None


def group_files_by_identifier(filenames: List[str]) -> Dict[str, List[str]]:
    groups = {}
    for fname in filenames:
        if fname.lower().endswith('.xlsx'):
            identifier = extract_identifier(fname)
            if identifier:
                groups.setdefault(identifier, []).append(fname)
            else:
                logging.warning(
                    f"Filename does not match expected format: {fname}")
    return groups


def save_dataframe_to_excel(df: pd.DataFrame, output_path: str, row_limit: int = ROW_LIMIT_PER_SHEET) -> None:
    n_rows = len(df)
    if n_rows == 0:
        logging.warning(f"No data to save for {output_path}. Skipping file.")
        return
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for i in range((n_rows + row_limit - 1) // row_limit):
                start, end = i * row_limit, min((i + 1) * row_limit, n_rows)
                sheet_name = f"Sheet{i+1}"
                df.iloc[start:end].to_excel(
                    writer, index=False, sheet_name=sheet_name)
            logging.info(
                f"Saved to {output_path} with {((n_rows + row_limit - 1) // row_limit)} sheet(s).")
    except Exception as e:
        logging.error(f"Failed to save Excel file '{output_path}': {e}")


def merge_excels(input_dir: str = 'final', output_dir: str = 'merged', row_limit: int = ROW_LIMIT_PER_SHEET) -> None:
    """
    Groups, merges, and writes Excel files from input_dir to output_dir.
    Each group (by identifier) is saved to one or more Excel sheets (if >1,024,000 rows).
    """
    if not os.path.exists(input_dir):
        logging.error(f"Input dir '{input_dir}' does not exist! Exiting.")
        return

    os.makedirs(output_dir, exist_ok=True)

    try:
        all_files = os.listdir(input_dir)
    except Exception as e:
        logging.error(f"Cannot list files in '{input_dir}': {e}")
        return

    groups = group_files_by_identifier(all_files)
    if not groups:
        logging.info(
            "No matching Excel files by identifier found. Nothing to merge.")
        return

    for identifier, filelist in groups.items():
        merged_rows = []
        for fname in filelist:
            full_path = os.path.join(input_dir, fname)
            try:
                df = pd.read_excel(full_path, engine='openpyxl')
                df['__sourcefile__'] = fname
                merged_rows.append(df)
            except Exception as e:
                logging.error(f"Could not read '{fname}': {e}")
        if merged_rows:
            try:
                big_df = pd.concat(merged_rows, ignore_index=True)
                output_file = os.path.join(
                    output_dir, f"merged_{identifier}.xlsx")
                save_dataframe_to_excel(big_df, output_file, row_limit)
            except Exception as e:
                logging.error(
                    f"Error during merging/writing for '{identifier}': {e}")
        else:
            logging.warning(
                f"Nothing to merge for group {identifier}. Skipping output.")
