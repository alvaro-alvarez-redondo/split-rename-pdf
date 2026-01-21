import os
import sys
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
import re

# ---------------------------------------------------------------------
# directories
# ---------------------------------------------------------------------

base_dir = os.path.abspath(os.path.dirname(__file__))
output_folder = os.path.join(base_dir, "output")

print("\033[1;34mstarting script...\033[0m")  # blue for info

if not os.path.exists(output_folder):
    os.makedirs(output_folder)
    print(f"📂 folder \033[1;34m'output'\033[0m created")

# ---------------------------------------------------------------------
# editable configuration
# ---------------------------------------------------------------------

excel_filename = os.path.join(base_dir, "rename-pdf-mapping.xlsx")
output_pattern = "{yearbook}_{year}_{category}_{first_page}_{last_page}_{product}"

required_columns = [
    "yearbook", "year", "category", "products",
    "yearbook_start", "yearbook_end", "pdf_start", "pdf_end"
]

# ---------------------------------------------------------------------
# helper functions
# ---------------------------------------------------------------------

def sanitize_filename(name: str) -> str:
    """
    clean a string to use as filename:
    - replace invalid chars with '_'
    - replace spaces with '_'
    - remove duplicates at start/end
    - lowercase everything
    """
    cleaned = re.sub(r'[\/:*?"<>|\s]+', '_', name.strip())
    return cleaned.strip('_').lower()

# ---------------------------------------------------------------------
# main function
# ---------------------------------------------------------------------

def split_and_rename_pdf():
    # locate pdf
    pdf_files = [f for f in os.listdir(base_dir) if f.lower().endswith(".pdf")]

    if len(pdf_files) == 0:
        print("\033[1;31m❌ error: no pdf found in the script folder.\033[0m")
        sys.exit(1)
    elif len(pdf_files) > 1:
        print("\033[1;31m❌ error: more than one pdf found in the script folder.\033[0m")
        sys.exit(1)

    input_pdf = pdf_files[0]
    input_pdf_path = os.path.join(base_dir, input_pdf)

    # locate or create excel
    if not os.path.exists(excel_filename):
        print(f"\033[1;33m⚠️  excel file 'rename-pdf-mapping.xlsx' not found. a template will be created.\033[0m")
        df_template = pd.DataFrame(columns=required_columns)
        df_template.to_excel(excel_filename, index=False)
        print(f"\033[1;33m📄 excel file created. fill it out and run the script again.\033[0m")
        sys.exit(0)

    # read excel
    df = pd.read_excel(excel_filename)
    missing_columns = set(required_columns) - set(df.columns)
    if missing_columns:
        print(f"\033[1;31m❌ error: missing columns in excel: {missing_columns}\033[0m")
        sys.exit(1)
    if df.empty:
        print(f"\033[1;31m❌ error: excel file is empty.\033[0m")
        sys.exit(1)

    # read pdf
    reader = PdfReader(input_pdf_path)
    total_pages = len(reader.pages)
    total_outputs = len(df)

    # ---------------------------------------------
    # check for existing output files
    # ---------------------------------------------
    existing_files = []

    for index, row in df.iterrows():
        fields = ["yearbook", "year", "category", "products"]
        data = {field: sanitize_filename(str(row[field])) for field in fields}

        yearbook = data["yearbook"]
        year = data["year"]
        category = data["category"]
        product = data["products"]

        mapped_first = int(row["yearbook_start"])
        mapped_last = int(row["yearbook_end"])

        new_base_name = output_pattern.format(
            yearbook=yearbook,
            year=year,
            category=category,
            first_page=mapped_first,
            last_page=mapped_last,
            product=product,
        )

        file_path = os.path.join(output_folder, f"{new_base_name}.pdf")
        if os.path.exists(file_path):
            existing_files.append(file_path)

    overwrite_all = False
    num_existing = len(existing_files)
    if num_existing > 0:
        while True:
            choice = input(f"\033[1;33m⚠️  {num_existing} file(s) already exist, overwrite them all? (y/n): \033[0m").strip().lower()
            if choice == 'y':
                overwrite_all = True
                break
            elif choice == 'n':
                overwrite_all = False
                break
            else:
                print(f"\033[1;33m⚠️  please enter 'y' or 'n'..\033[0m")

    # progress bar colors
    progress_color = "\033[1;36m"  # cyan
    reset_color = "\033[0m"

    # main loop
    for index, row in df.iterrows():
        try:
            # extract data and sanitize names
            fields = ["yearbook", "year", "category", "products"]
            data = {field: sanitize_filename(str(row[field])) for field in fields}

            yearbook = data["yearbook"]
            year = data["year"]
            category = data["category"]
            product = data["products"]

            mapped_first = int(row["yearbook_start"])
            mapped_last = int(row["yearbook_end"])
            pdf_first = int(row["pdf_start"])
            pdf_last = int(row["pdf_end"])
        except Exception as e:
            print(f"\033[1;31m❌ error reading row {index + 1}: {e}\033[0m")
            sys.exit(1)

        # validate page ranges
        if pdf_first < 1 or pdf_last > total_pages or pdf_first > pdf_last:
            print(f"\033[1;31m❌ invalid page range for '{product}': {pdf_first}-{pdf_last} (pdf has {total_pages} pages)\033[0m")
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

        final_output_path = os.path.join(output_folder, f"{new_base_name}.pdf")

        # handle overwriting or generate unique name
        if not overwrite_all and os.path.exists(final_output_path):
            counter = 1
            base_name = new_base_name
            while os.path.exists(final_output_path):
                final_output_path = os.path.join(output_folder, f"{base_name}_{counter}.pdf")
                counter += 1

        with open(final_output_path, "wb") as f:
            writer.write(f)

        # progress bar
        progress = (index + 1) / total_outputs
        bar_length = 30
        filled_length = int(bar_length * progress)
        bar = "█" * filled_length + "-" * (bar_length - filled_length)
        print(f"\r{progress_color}progress: |{bar}| {index + 1}/{total_outputs} (pages {pdf_first}-{pdf_last}){reset_color}", end="")

    print(f"\n\033[1;32m✅ all pdfs were successfully extracted and renamed.\033[0m")

# ---------------------------------------------------------------------
# entry point
# ---------------------------------------------------------------------

if __name__ == "__main__":
    try:
        split_and_rename_pdf()
    except Exception as e:
        print(f"\033[1;31m❌ unexpected error occurred: {e}\033[0m")
        sys.exit(1)
