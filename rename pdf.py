import os
import sys
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
import re

# ---------------------------------------------------------------------
# Directories
# ---------------------------------------------------------------------

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
OUTPUT_FOLDER = os.path.join(BASE_DIR, "output")

print("\033[1;34mStarting script...\033[0m")  # Blue for info

# Create output folder if it doesn't exist
if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)
    print(f"📂 Folder \033[1;34m'output'\033[0m created")

# ---------------------------------------------------------------------
# Editable Configuration
# ---------------------------------------------------------------------

EXCEL_FILENAME = os.path.join(BASE_DIR, "rename pdf mapping.xlsx")
output_pattern = "{yearbook}_{year}_{category}_{first_page}_{last_page}_{product}"

REQUIRED_COLUMNS = [
    "Yearbook", "Year", "Category", "Products",
    "Yearbook start", "Yearbook end", "PDF start", "PDF end"
]

# ---------------------------------------------------------------------
# Helper Functions
# ---------------------------------------------------------------------

def sanitize_filename(name: str) -> str:
    """Remove invalid characters from filenames."""
    return re.sub(r'[\/:*?"<>|]', '_', name.strip())

# ---------------------------------------------------------------------
# Main Function
# ---------------------------------------------------------------------

def split_and_rename_pdf():
    # -------------------------------------------------
    # Locate PDF
    # -------------------------------------------------
    pdf_files = [f for f in os.listdir(BASE_DIR) if f.lower().endswith(".pdf")]

    if len(pdf_files) == 0:
        print("\033[1;31m❌ Error: No PDF found in the script folder.\033[0m")
        sys.exit(1)
    elif len(pdf_files) > 1:
        print("\033[1;31m❌ Error: More than one PDF found in the script folder.\033[0m")
        sys.exit(1)

    input_pdf = pdf_files[0]
    input_pdf_path = os.path.join(BASE_DIR, input_pdf)

    # -------------------------------------------------
    # Locate or create Excel
    # -------------------------------------------------
    if not os.path.exists(EXCEL_FILENAME):
        print(f"\033[1;33m⚠️  Excel file 'rename pdf mapping.xlsx' not found. A template will be created.\033[0m")
        df_template = pd.DataFrame(columns=REQUIRED_COLUMNS)
        df_template.to_excel(EXCEL_FILENAME, index=False)
        print(f"\033[1;33m📄 Excel file created. Fill it out and run the script again.\033[0m")
        sys.exit(0)

    # -------------------------------------------------
    # Read Excel
    # -------------------------------------------------
    df = pd.read_excel(EXCEL_FILENAME)
    missing_columns = set(REQUIRED_COLUMNS) - set(df.columns)
    if missing_columns:
        print(f"\033[1;31m❌ Error: Missing columns in Excel: {missing_columns}\033[0m")
        sys.exit(1)
    if df.empty:
        print("\033[1;31m❌ Error: Excel file is empty.\033[0m")
        sys.exit(1)

    # -------------------------------------------------
    # Read PDF
    # -------------------------------------------------
    reader = PdfReader(input_pdf_path)
    total_pages = len(reader.pages)
    total_outputs = len(df)

    # Progress bar colors
    progress_color = "\033[1;36m"  # Cyan
    reset_color = "\033[0m"

    # -------------------------------------------------
    # Main loop
    # -------------------------------------------------
    for index, row in df.iterrows():
        try:
            # Extract data and sanitize names
            yearbook = sanitize_filename(str(row["Yearbook"]))
            year = sanitize_filename(str(row["Year"]))
            category = sanitize_filename(str(row["Category"]))
            product = sanitize_filename(str(row["Products"]))

            mapped_first = int(row["Yearbook start"])
            mapped_last = int(row["Yearbook end"])
            pdf_first = int(row["PDF start"])
            pdf_last = int(row["PDF end"])
        except Exception as e:
            print(f"\033[1;31m❌ Error reading row {index + 1}: {e}\033[0m")
            sys.exit(1)

        # Validate page ranges
        if pdf_first < 1 or pdf_last > total_pages or pdf_first > pdf_last:
            print(f"\033[1;31m❌ Invalid page range for '{product}': {pdf_first}-{pdf_last} (PDF has {total_pages} pages)\033[0m")
            sys.exit(1)

        writer = PdfWriter()
        for page_num in range(pdf_first - 1, pdf_last):
            writer.add_page(reader.pages[page_num])

        new_base_name = output_pattern.format(
            yearbook=yearbook,
            year=year,
            category=category,
            first_page=mapped_first,
            last_page=mapped_last,
            product=product,
        )

        output_path = os.path.join(OUTPUT_FOLDER, f"{new_base_name}.pdf")

        # Avoid overwriting
        counter = 1
        final_output_path = output_path
        while os.path.exists(final_output_path):
            final_output_path = os.path.join(OUTPUT_FOLDER, f"{new_base_name}_{counter}.pdf")
            counter += 1

        with open(final_output_path, "wb") as f:
            writer.write(f)

        # Progress bar
        progress = (index + 1) / total_outputs
        bar_length = 30
        filled_length = int(bar_length * progress)
        bar = "█" * filled_length + "-" * (bar_length - filled_length)
        print(f"\r{progress_color}Progress: |{bar}| {index + 1}/{total_outputs} (pages {pdf_first}-{pdf_last}){reset_color}", end="")

    print(f"\n\033[1;32m✅ All PDFs were successfully extracted and renamed.\033[0m")

# ---------------------------------------------------------------------
# Entry Point
# ---------------------------------------------------------------------

if __name__ == "__main__":
    try:
        split_and_rename_pdf()
    except Exception as e:
        print(f"\033[1;31m❌ Unexpected error occurred: {e}\033[0m")
        sys.exit(1)
