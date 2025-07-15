import pandas as pd
from pathlib import Path
import re


def apply_final_transformations(input_dir="transformed", output_dir="final"):
    """
    Process single-sheet Excel files for final transformations

    Parameters:
    input_dir (str): Source directory with transformed files
    output_dir (str): Target directory for final output
    """
    # Configuration
    # date_pattern = re.compile(r'\b\d{2}-\d{2}-\d{4}\b')
    min_fields = 7  # perno + Name + dob + sex + appid_receipt_no + address_parts
    date_pattern = re.compile(
        r'\b(\d{2}[-\/]\d{2}[-\/]\d{4})\b')  # Enhanced pattern

    Path(output_dir).mkdir(parents=True, exist_ok=True)

    for excel_file in Path(input_dir).glob("*.xlsx"):
        processed_data = []
        # error_log = []
        output_path = Path(output_dir) / f"final_{excel_file.name}"
        try:
            # df = pd.read_excel(excel_file, header=None,
            #                    engine='openpyxl').drop(columns=[0, 1])
            df = pd.read_excel(excel_file, header=None, skiprows=1,
                               engine='openpyxl').dropna(how='all').drop(columns=[0, 1])

            for idx, row in df.iterrows():
                raw_text = ''
                try:
                    # Extract and clean all text from row
                    raw_text = ' '.join([
                        c.strftime('%d-%m-%Y') if isinstance(c, pd.Timestamp)
                        else str(c).strip().replace('/', '-')
                        for c in row if pd.notna(c)
                    ])

                    # Split into components while preserving multi-word fields
                    tokens = raw_text.split()
                    if len(tokens) < min_fields:
                        raise ValueError(
                            f"Only {len(tokens)} components found (minimum {min_fields} required)")

                    # Find date position with priority scanning
                    date_index = next((i for i, t in enumerate(
                        tokens) if date_pattern.match(t)), None)

                    # Validate critical components
                    if not date_index or date_index < 2:
                        raise ValueError(
                            f"Date not found in expected position. Tokens: {tokens[:6]}")
                    if len(tokens) < date_index + 4:
                        raise ValueError(
                            f"Missing components after date. Found: {tokens[date_index:]}")

                    # Extract components with dynamic positioning
                    components = {
                        'voter_id': tokens[0],
                        'name_parts': tokens[1:date_index],
                        'dob': tokens[date_index],
                        'gender': tokens[date_index+1],
                        'reg_id': tokens[date_index+2],
                        'address_parts': tokens[date_index+3:]
                    }

                    # Validation checks
                    if not components['voter_id'].isdigit():
                        raise ValueError(
                            f"Invalid Voter ID format: {components['voter_id']}")
                    if components['gender'].upper() not in {'M', 'F'}:
                        raise ValueError(
                            f"Invalid gender: {components['gender']}")
                    if not components['address_parts']:
                        raise ValueError("Missing address information")

                    processed_data.append({
                        'perno': components['voter_id'],
                        'surname': components['name_parts'][0] if components['name_parts'] else '',
                        'othernames': ' '.join(components['name_parts'][1:]) if len(components['name_parts']) > 1 else '',
                        'dob': components['dob'],
                        'sex': components['gender'].upper(),
                        'appid_receipt_no': components['reg_id'],
                        'village': ' '.join(components['address_parts'])
                    })

                except Exception as e:
                    print(f"Critical system error processing excel: {str(e)}")

            # Save outputs
            pd.DataFrame(processed_data).to_excel(output_path, index=False)
            # if error_log:
            #     pd.DataFrame(error_log).to_csv(output_path, index=False)

        except Exception as e:
            print(f"Critical system error reading excel: {str(e)}")
            return False
