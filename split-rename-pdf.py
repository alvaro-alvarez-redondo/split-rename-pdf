import sys
import re
from pathlib import Path

import pandas as pd
from PyPDF2 import PdfReader, PdfWriter

# ---------------------------------------------------------------------
# ANSI colors
# ---------------------------------------------------------------------

ERR = "\033[1;31m"     # red bold
WARN = "\033[1;33m"    # yellow bold
INFO = "\033[1;34m"    # blue bold
HELP = "\033[1;37m"    # white bold
OK = "\033[1;32m"      # green bold
RST = "\033[0m"

# ---------------------------------------------------------------------
# Editable configuration
# ---------------------------------------------------------------------

BASE_DIR = Path(__file__).parent.resolve()
EXCEL_FILENAME = BASE_DIR / "rename-pdf-mapping.xlsx"
OUTPUT_PATTERN = "{yearbook}_{category}_{year}_{first_page}_{last_page}_{product}"

REQUIRED_COLUMNS = [
    "yearbook", "year", "category", "products",
    "yearbook_start", "yearbook_end", "pdf_start", "pdf_end"
]

FIELDS_TO_SANITIZE = ["yearbook", "year", "category", "products"]

# ---------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------

def print_error(message: str, help_lines: list[str]):
    """Print error message with optional help tips in colored format."""
    print()

    print(f"{ERR}{message}{RST}")
    for line in help_lines:
        print(f"{HELP}• {line}{RST}")
    print()


def sanitize_filename(name: str) -> str:
    """Sanitize a string to be a safe file name."""
    cleaned = re.sub(r'[\/:*?"<>|\s]+', '_', name.strip())
    return cleaned.strip('_').lower()


def check_pdf_files(base_dir: Path) -> Path:
    """Ensure exactly one PDF file exists in the directory."""
    pdf_files = list(base_dir.glob("*.pdf"))
    if len(pdf_files) == 0:
        print_error(
            "No PDF found in the script folder.",
            [
                "Place exactly one PDF file in the same folder as the script",
                "Make sure the file has a .pdf extension"
            ]
        )
        sys.exit(1)
    elif len(pdf_files) > 1:
        print_error(
            "More than one PDF found in the script folder.",
            [
                "Leave only one PDF file in the folder",
                "Move or delete extra PDF files before running the script"
            ]
        )
        sys.exit(1)
    return pdf_files[0]


def create_output_folder(base_dir: Path, pdf_name: str) -> Path:
    """Create an output folder for split PDFs."""
    folder = base_dir / pdf_name.stem
    folder.mkdir(exist_ok=True)
    print(f"Folder {INFO}'{folder.name}'{RST} ready")
    return folder


def load_excel(excel_path: Path) -> pd.DataFrame:
    """Load Excel file and validate columns."""
    if not excel_path.exists():
        print(f"{WARN}Excel file '{excel_path.name}' not found. A template will be created.{RST}")
        pd.DataFrame(columns=REQUIRED_COLUMNS).to_excel(excel_path, index=False)
        print(f"{WARN}Excel file created. Fill it out and run the script again.{RST}")
        sys.exit(0)

    df = pd.read_excel(excel_path)
    missing_cols = set(REQUIRED_COLUMNS) - set(df.columns)
    if missing_cols:
        print_error(
            f"Missing columns in excel: {missing_cols}",
            [
                "Open the Excel file and verify all required columns exist",
                "Do not rename or remove any required column headers"
            ]
        )
        sys.exit(1)
    if df.empty:
        print_error(
            "Excel file is empty.",
            [
                "Add at least one data row below the header",
                "Fill in all required columns before running the script"
            ]
        )
        sys.exit(1)
    return df


def generate_output_name(row) -> str:
    """Generate sanitized PDF output name from row data."""
    data = {field: sanitize_filename(str(getattr(row, field))) for field in FIELDS_TO_SANITIZE}
    return OUTPUT_PATTERN.format(
        yearbook=data["yearbook"],
        category=data["category"],
        year=data["year"],
        first_page=int(row.yearbook_start),
        last_page=int(row.yearbook_end),
        product=data["products"]
    )


def extract_pdf_pages(reader: PdfReader, start: int, end: int, output_path: Path):
    """Extract pages from reader and write to output_path."""
    writer = PdfWriter()
    for i in range(start - 1, end):
        writer.add_page(reader.pages[i])
    with open(output_path, "wb") as f:
        writer.write(f)


# ---------------------------------------------------------------------
# Main function
# ---------------------------------------------------------------------

def split_and_rename_pdf():
    print(f"{INFO}Starting script...{RST}")

    # Locate PDF
    input_pdf_path = check_pdf_files(BASE_DIR)
    output_folder = create_output_folder(BASE_DIR, input_pdf_path)

    # Load Excel
    df = load_excel(EXCEL_FILENAME)
    total_pages = len(PdfReader(input_pdf_path).pages)
    total_outputs = len(df)

    # Check for existing files
    existing_files = {
        generate_output_name(row)
        for row in df.itertuples()
        if (output_folder / f"{generate_output_name(row)}.pdf").exists()
    }

    overwrite_all = False
    if existing_files:
        while True:
            choice = input(
                f"{WARN}{len(existing_files)} file(s) already exist, overwrite them all? (y/n): {RST}"
            ).strip().lower()
            if choice in ("y", "n"):
                overwrite_all = choice == "y"
                break
            print(f"{WARN}Please enter 'y' or 'n'.{RST}")

    # Process PDFs
    reader = PdfReader(input_pdf_path)
    for idx, row in enumerate(df.itertuples(), start=1):
        try:
            pdf_first, pdf_last = int(row.pdf_start), int(row.pdf_end)
            yearbook_first, yearbook_last = int(row.yearbook_start), int(row.yearbook_end)
            output_name = generate_output_name(row)
        except Exception as e:
            print_error(f"Error reading row {idx}: {e}", [
                "Check the indicated row in the Excel file",
                "Ensure numeric columns contain only numbers",
                "Avoid empty cells in required fields"
            ])
            sys.exit(1)

        # Validate page range
        if pdf_first < 1 or pdf_last > total_pages or pdf_first > pdf_last:
            print()
            print_error(f"Invalid page range for '{row.products}': {pdf_first}-{pdf_last} (pdf has {total_pages} pages)", [
                "Verify pdf_start and pdf_end values",
                "The range must be within the total number of PDF pages",
                "pdf_start must not be greater than pdf_end"
            ])
            sys.exit(1)

        final_output_path = output_folder / f"{output_name}.pdf"
        if not overwrite_all and final_output_path.exists():
            counter = 1
            base_name = output_folder / output_name
            while final_output_path.exists():
                final_output_path = output_folder / f"{output_name}_{counter}.pdf"
                counter += 1

        extract_pdf_pages(reader, pdf_first, pdf_last, final_output_path)

        # Progress bar
        progress = idx / total_outputs
        bar_length = 30
        filled_length = int(bar_length * progress)
        bar = "█" * filled_length + "-" * (bar_length - filled_length)
        print(f"\r{INFO}Progress: |{bar}| {idx}/{total_outputs} (pages {pdf_first}-{pdf_last}){RST}", end="")

    print(f"\n{OK}All PDFs were successfully extracted and renamed.{RST}")


# ---------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------

if __name__ == "__main__":
    try:
        split_and_rename_pdf()
    except SystemExit:
        pass
    except Exception as e:
        print_error(f"Unexpected error occurred: {e}", [
            "Make sure the PDF is not open in another program",
            "Verify write permissions in the output folder",
            "Try running the script again"
        ])
        sys.exit(1)
