from pathlib import Path
import pandas as pd
from openpyxl.utils import get_column_letter  # Added missing import

def transform_excels(input_dir="converted", output_dir="transformed"):
    Path(output_dir).mkdir(parents=True, exist_ok=True)

    for excel_file in Path(input_dir).glob("*.xlsx"):
        output_path = Path(output_dir) / f"transformed_{excel_file.name}"

        try:
            all_sheets = pd.read_excel(excel_file, engine='openpyxl', sheet_name=None)
            processed_data = []

            for sheet_name, df in all_sheets.items():
                try:
                    # Data cleaning process
                    df_clean = df.iloc[5:].reset_index(drop=True)
                    df_clean.columns = df_clean.iloc[0]
                    df_clean = df_clean.iloc[1:].dropna(how='all')

                    # Fixed: Replace deprecated applymap with vectorized operation
                    df_clean = df_clean.apply(lambda col: col.str.strip() if col.dtype == 'object' else col)

                    # Remove page rows
                    page_mask = df_clean.apply(lambda row: row.astype(str).str.contains('Page').any(), axis=1)
                    df_clean = df_clean[~page_mask]

                    if not df_clean.empty:
                        processed_data.append(df_clean)

                except Exception as e:
                    print(f"Sheet error in {sheet_name}: {str(e)}")

            if processed_data:
                combined_df = pd.concat(processed_data, ignore_index=True)
                combined_df.insert(0, 'Serial No', range(1, len(combined_df)+1))

                with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                    combined_df.to_excel(writer, index=False)
                    worksheet = writer.sheets['Sheet1']

                    # Column width adjustment (now with proper import)
                    for col_num, column_name in enumerate(combined_df.columns, 1):
                        max_length = max(
                            combined_df[column_name].astype(str).map(len).max(),
                            len(str(column_name))
                        ) + 2
                        worksheet.column_dimensions[get_column_letter(col_num)].width = max_length

                print(f"Successfully processed: {excel_file.name}")

        except Exception as e:
            print(f"Failed processing {excel_file.name}: {str(e)}")
transform_excels()
