import sys
import subprocess
from pathlib import Path
import re

# ---------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------
BASE_DIR = Path(__file__).parent.resolve()
EXCEL_FILENAME = BASE_DIR / "rename-pdf-mapping.xlsx"
OUTPUT_PATTERN = "{yearbook}_{category}_{year}_{first_page}_{last_page}_{product}"
REQUIRED_COLUMNS = [
    "yearbook", "year", "category", "products",
    "yearbook_start", "yearbook_end", "pdf_start", "pdf_end"
]
FIELDS_TO_SANITIZE = ["yearbook", "year", "category", "products"]
REQUIRED_PACKAGES = ["pandas", "PyPDF2", "openpyxl"]

# ---------------------------------------------------------------------
# Python version check
# ---------------------------------------------------------------------
if sys.version_info < (3, 8):
    print("\033[1;31mError: Python 3.8+ is required\033[0m")
    sys.exit(1)

# ---------------------------------------------------------------------
# Automatic requirements.txt creation
# ---------------------------------------------------------------------
REQUIREMENTS_FILE = BASE_DIR / "requirements.txt"
if not REQUIREMENTS_FILE.exists():
    REQUIREMENTS_FILE.write_text("\n".join(REQUIRED_PACKAGES))
    print("\033[1;33m'requirements.txt' created. The script will attempt to install missing packages automatically.\033[0m")
    sys.exit(0)

# ---------------------------------------------------------------------
# Automatic installation of missing packages
# ---------------------------------------------------------------------
missing_modules = []
for pkg in REQUIRED_PACKAGES:
    try:
        __import__(pkg)
    except ModuleNotFoundError:
        missing_modules.append(pkg)

if missing_modules:
    print(f"\033[1;33mInstalling missing packages: {', '.join(missing_modules)}...\033[0m")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", *missing_modules])
    except subprocess.CalledProcessError:
        print(f"\033[1;31mFailed to install packages: {', '.join(missing_modules)}\033[0m")
        sys.exit(1)

# ---------------------------------------------------------------------
# Now all packages are guaranteed to be installed → imports at the top
# ---------------------------------------------------------------------
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter

# ---------------------------------------------------------------------
# ANSI colors
# ---------------------------------------------------------------------
ERR = "\033[1;31m"
WARN = "\033[1;33m"
INFO = "\033[1;34m"
HELP = "\033[1;37m"
OK = "\033[1;32m"
RST = "\033[0m"

# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------
def print_error(message: str, help_lines: list[str]):
    print(f"\n{ERR}{message}{RST}")
    for line in help_lines:
        print(f"{HELP}• {line}{RST}")
    print()

def sanitize_filename(value: str) -> str:
    cleaned = re.sub(r'[\/:*?"<>|\s]+', '_', value.strip())
    return cleaned.strip('_').lower()

def ask_yes_no(prompt: str) -> bool:
    choice = input(prompt).strip().lower()
    if choice not in {"y", "n"}:
        print_error("Invalid input.", ["Please enter only 'y' or 'n'"])
        sys.exit(1)
    return choice == "y"

def check_pdf_files(base_dir: Path) -> Path:
    pdfs = list(base_dir.glob("*.pdf"))
    if len(pdfs) != 1:
        print_error(
            "Expected exactly one PDF file.",
            ["Place only one .pdf file in the script directory"]
        )
        sys.exit(1)
    return pdfs[0]

def create_output_folder(base_dir: Path, pdf_path: Path) -> Path:
    folder = base_dir / pdf_path.stem
    folder.mkdir(exist_ok=True)
    print(f"{INFO}Output folder '{folder.name}' ready{RST}")
    return folder

def load_excel(path: Path) -> pd.DataFrame:
    if not path.exists():
        pd.DataFrame(columns=REQUIRED_COLUMNS).to_excel(path, index=False)
        print_error(
            f"Excel file '{path.name}' was created.",
            ["Fill it with data and run the script again"]
        )
        sys.exit(0)

    df = pd.read_excel(
        path,
        dtype={
            "yearbook": "string",
            "year": "string",
            "category": "string",
            "products": "string",
            "yearbook_start": "Int64",
            "yearbook_end": "Int64",
            "pdf_start": "Int64",
            "pdf_end": "Int64",
        }
    )

    missing = set(REQUIRED_COLUMNS) - set(df.columns)
    if missing or df.empty:
        print_error(
            "Excel validation failed.",
            ["Verify required columns exist", "Ensure the file is not empty"]
        )
        sys.exit(1)

    return df

# ---------------------------------------------------------------------
# Data processing
# ---------------------------------------------------------------------
def handle_empty_products(df: pd.DataFrame) -> pd.DataFrame:
    empty = df["products"].isna() | (df["products"].astype(str).str.strip() == "")
    if not empty.any():
        return df

    intentional = ask_yes_no(
        f"{WARN}{empty.sum()} empty 'products'. Is this intentional? (y/n): {RST}"
    )

    if not intentional:
        print_error("Empty products detected.", ["Fill all 'products' cells in Excel and rerun"])
        sys.exit(1)

    df.loc[empty, "products"] = ""
    return df

def generate_output_name(row) -> str:
    data = {field: sanitize_filename(str(getattr(row, field))) for field in FIELDS_TO_SANITIZE}

    if data["products"] == "":
        return f"{data['yearbook']}_{data['category']}_{data['year']}_{int(row.yearbook_start)}_{int(row.yearbook_end)}"

    return OUTPUT_PATTERN.format(
        yearbook=data["yearbook"],
        category=data["category"],
        year=data["year"],
        first_page=int(row.yearbook_start),
        last_page=int(row.yearbook_end),
        product=data["products"]
    )

def unique_output_path(folder: Path, name: str) -> Path:
    candidates = (folder / f"{name}.pdf", *(folder / f"{name}_{i}.pdf" for i in range(1, 10_000)))
    return next(p for p in candidates if not p.exists())

def extract_pdf_pages(reader: PdfReader, start: int, end: int, output: Path):
    writer = PdfWriter()
    for i in range(start - 1, end):
        writer.add_page(reader.pages[i])
    with open(output, "wb") as f:
        writer.write(f)

# ---------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------
def split_and_rename_pdf():
    print(f"{INFO}Starting...{RST}")

    pdf_path = check_pdf_files(BASE_DIR)
    output_folder = create_output_folder(BASE_DIR, pdf_path)

    df = load_excel(EXCEL_FILENAME)
    df = handle_empty_products(df)

    reader = PdfReader(pdf_path)
    total_pages = len(reader.pages)

    df["output_name"] = df.apply(generate_output_name, axis=1)

    existing = df["output_name"].map(lambda n: (output_folder / f"{n}.pdf").exists())

    overwrite_all = False
    if existing.any():
        overwrite_all = ask_yes_no(
            f"{WARN}{existing.sum()} files already exist. Overwrite all? (y/n): {RST}"
        )

    total = len(df)

    for idx, row in enumerate(df.itertuples(), start=1):
        pdf_start, pdf_end = int(row.pdf_start), int(row.pdf_end)

        if pdf_start < 1 or pdf_end > total_pages or pdf_start > pdf_end:
            print_error(f"Invalid page range in row {idx}.", ["Check pdf_start and pdf_end values"])
            sys.exit(1)

        output_path = output_folder / f"{row.output_name}.pdf"

        if output_path.exists() and not overwrite_all:
            output_path = unique_output_path(output_folder, row.output_name)

        extract_pdf_pages(reader, pdf_start, pdf_end, output_path)

        progress = idx / total
        bar = "█" * int(30 * progress) + "-" * (30 - int(30 * progress))
        print(f"\r{INFO}|{bar}| {idx}/{total}{RST}", end="")

    print(f"\n{OK}PDFs successfully created.{RST}")

# ---------------------------------------------------------------------
if __name__ == "__main__":
    try:
        split_and_rename_pdf()
    except Exception as e:
        print_error(
            f"Unexpected error: {e}",
            ["Ensure the PDF is closed", "Check write permissions"]
        )
        sys.exit(1)
