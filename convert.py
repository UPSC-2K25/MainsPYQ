import json
import os
import re
import pandas as pd


def parse_question_cell(cell_value):
    """Parses the question cell.

    If multiple sub-questions (e.g., starting with a., b., c., etc.) are entered
    with newlines, returns them as a structured list of strings. If it's a
    single question, returns a single clean string. If empty/null, returns None.
    """
    if pd.isna(cell_value):
        return None

    text = str(cell_value).strip()
    if not text:
        return None

    # Split lines by carriage return / line feed
    raw_lines = [
        line.strip() for line in re.split(r"\r?\n+", text) if line.strip()
    ]

    if not raw_lines:
        return None

    if len(raw_lines) == 1:
        return raw_lines[0]

    # Multiple sub-questions present in single cell
    return raw_lines


def clean_row_value(val):
    """Sanitizes cell values:

    - Converts pandas NaN / blank text to None
    - Converts float integers (e.g., 2024.0 -> 2024)
    - Strips whitespace from strings
    """
    if pd.isna(val):
        return None

    if isinstance(val, float) and val.is_integer():
        return int(val)

    if isinstance(val, str):
        cleaned = val.strip()
        return cleaned if cleaned else None

    return val


def convert_workbook_to_json(
    excel_file_path, output_dir=".", drop_null_keys=False
):
    """Reads all sheets from the Excel file and generates individual JSON files.

    - drop_null_keys: If True, completely excludes keys with null values from JSON.
                      If False, sets them explicitly to null (recommended for schema consistency).
    """
    if not os.path.exists(excel_file_path):
        print(f"Error: File '{excel_file_path}' not found.")
        return

    # Load all sheets: {sheet_name: DataFrame}
    workbook = pd.read_excel(excel_file_path, sheet_name=None)

    expected_columns = [
        "year",
        "paper",
        "sec",
        "qno",
        "question",
        "marks_words_time",
    ]

    total_sheets = len(workbook)
    print(
        f"Found {total_sheets} sheets in '{excel_file_path}'. Processing...\n"
    )

    for sheet_name, df in workbook.items():
        # Drop completely ghost/blank rows often left at the bottom of Excel sheets
        df = df.dropna(how="all")

        # Normalize header row: strip whitespace and lowercase
        df.columns = [
            str(col).strip().lower().replace(" ", "_") for col in df.columns
        ]

        sheet_records = []

        for _, row in df.iterrows():
            record = {}
            for col in expected_columns:
                val = row.get(col, None)

                if col == "question":
                    cleaned_val = parse_question_cell(val)
                else:
                    cleaned_val = clean_row_value(val)

                if drop_null_keys:
                    if cleaned_val is not None:
                        record[col] = cleaned_val
                else:
                    record[col] = cleaned_val

            # Only append record if it contains at least one non-null field
            if any(v is not None for v in record.values()):
                sheet_records.append(record)

        # Sanitize sheet name for valid Windows file naming (avoid \ / : * ? " < > |)
        safe_sheet_name = re.sub(r'[\\/*?:"<>|]', "_", sheet_name).strip()
        output_filename = os.path.join(output_dir, f"{safe_sheet_name}.json")

        with open(output_filename, "w", encoding="utf-8") as f:
            json.dump(sheet_records, f, indent=2, ensure_ascii=False)

        print(
            f"✔ Generated: '{output_filename}' ({len(sheet_records)} questions processed)"
        )

    print("\nAll sheets converted successfully.")


if __name__ == "__main__":
    # Specify the name of your Excel file
    EXCEL_FILE = "questions_bank.xlsx"

    # Set drop_null_keys=True if you do NOT want "sec": null in the output
    convert_workbook_to_json(EXCEL_FILE, drop_null_keys=False)