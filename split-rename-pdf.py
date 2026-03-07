from __future__ import annotations
import sys
import subprocess
from pathlib import Path
import re

# ---------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------
BASE_DIR = Path(__file__).parent.resolve()
EXCEL_FILENAME = BASE_DIR / "split-rename-pdf-mapping.xlsx"
OUTPUT_PATTERN = "{yearbook}_{category}_{year}_{first_page}_{last_page}_{product}"
REQUIRED_COLUMNS = [
    "yearbook", "year", "category", "products",
    "yearbook_start", "yearbook_end", "pdf_start", "pdf_end"
]
FIELDS_TO_SANITIZE = ["yearbook", "year", "category", "products"]
REQUIRED_PACKAGES = ["pandas", "PyPDF2", "openpyxl"]

REQUIREMENTS_FILE = BASE_DIR / "requirements.txt"

# ---------------------------------------------------------------------
# ANSI colors
# ---------------------------------------------------------------------
ERR = "\033[1;31m"
WARN = "\033[1;33m"
INFO = "\033[1;34m"
HELP = "\033[1;37m"
OK = "\033[1;32m"
RST = "\033[0m"


class ControlledExit(Exception):
    """Raised for controlled, user-facing exits without SystemExit noise."""


def is_module_available(module_name: str) -> bool:
    """Return True when the module can be imported."""
    try:
        __import__(module_name)
    except ModuleNotFoundError:
        return False
    return True


def bootstrap_environment() -> bool:
    """Prepare runtime dependencies. Returns False when script should stop early."""
    if sys.version_info < (3, 8):
        print("\033[1;31mError: Python 3.8+ is required\033[0m")
        return False

    if not REQUIREMENTS_FILE.exists():
        REQUIREMENTS_FILE.write_text("\n".join(REQUIRED_PACKAGES))
        print("\033[1;33m'requirements.txt' created. The script will attempt to install missing packages automatically.\033[0m")

    missing_modules = list(filter(lambda pkg: not is_module_available(pkg), REQUIRED_PACKAGES))

    if missing_modules:
        print(f"\033[1;33mInstalling missing packages: {', '.join(missing_modules)}...\033[0m")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", *missing_modules])
        except subprocess.CalledProcessError:
            print(f"\033[1;31mFailed to install packages: {', '.join(missing_modules)}\033[0m")
            return False

    global pd, PdfReader, PdfWriter
    import pandas as pd_module
    from PyPDF2 import PdfReader as ReaderModule, PdfWriter as WriterModule

    pd = pd_module
    PdfReader = ReaderModule
    PdfWriter = WriterModule
    return True


# ---------------------------------------------------------------------
# Runtime import placeholders
# ---------------------------------------------------------------------
pd = None
PdfReader = None
PdfWriter = None


# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------
def print_error(message: str, help_lines: list[str]):
    formatted_help = "\n".join(map(lambda line: f"{HELP}• {line}{RST}", help_lines))
    print(f"{ERR}{message}{RST}\n{formatted_help}")


def sanitize_filename(value: str) -> str:
    cleaned = re.sub(r'[\/:*?"<>|\s]+', '_', value.strip())
    return cleaned.strip('_').lower()


def ask_yes_no(prompt: str) -> bool:
    choice = input(prompt).strip().lower()
    if choice not in {"y", "n"}:
        print_error("Invalid input.", ["Please enter only 'y' or 'n'"])
        raise ControlledExit
    return choice == "y"


def check_pdf_files(base_dir: Path) -> Path:
    pdfs = list(base_dir.glob("*.pdf"))
    if len(pdfs) != 1:
        print_error(
            "Expected exactly one PDF file.",
            ["Place only one .pdf file in the script directory"]
        )
        raise ControlledExit
    return pdfs[0]


def create_output_folder(base_dir: Path, pdf_path: Path) -> Path:
    folder = base_dir / pdf_path.stem
    folder.mkdir(exist_ok=True)
    print(f"{INFO}Output folder '{folder.name}' ready{RST}")
    return folder


# ---------------------------------------------------------------------
# Excel auto-detection & renaming
# ---------------------------------------------------------------------
def is_valid_mapping_excel(excel_path: Path) -> bool:
    """Check whether an Excel file has all required mapping columns."""
    try:
        columns = pd.read_excel(excel_path, nrows=0).columns
    except Exception:
        return False
    return set(REQUIRED_COLUMNS).issubset(columns)


def find_and_rename_valid_excel(target_path: Path) -> bool:
    candidates = list(filter(lambda path_item: path_item != target_path, target_path.parent.glob("*.xlsx")))
    valid_files = list(filter(is_valid_mapping_excel, candidates))


    if len(valid_files) == 1:
        valid_files[0].rename(target_path)
        print(f"{INFO}Excel file found and renamed as '{target_path.name}'{RST}")
        return True

    if len(valid_files) > 1:
        print_error(
            "Multiple valid Excel files found.",
            [
                "Leave only one Excel with the required columns",
                f"Expected name: {target_path.name}"
            ]
        )
        raise ControlledExit

    return False


def ensure_excel_exists(path: Path) -> bool:
    """Ensure mapping file exists. Returns True if created during this run."""
    if path.exists():
        return False

    renamed = find_and_rename_valid_excel(path)
    if renamed:
        return False

    pd.DataFrame(columns=REQUIRED_COLUMNS).to_excel(path, index=False)
    print_error(
        f"Excel file '{path.name}' was created.",
        ["Fill it with data and run the script again"]
    )
    return True


def load_excel(path: Path) -> pd.DataFrame:
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
        raise ControlledExit

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
        raise ControlledExit

    df.loc[empty, "products"] = ""
    return df


def generate_output_name(row) -> str:
    data = dict(map(lambda field: (field, sanitize_filename(str(getattr(row, field)))), FIELDS_TO_SANITIZE))

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


def unique_output_path(folder: Path, name: str, suffix: int = 0) -> Path:
    candidate = folder / (f"{name}.pdf" if suffix == 0 else f"{name}_{suffix}.pdf")
    if not candidate.exists():
        return candidate
    if suffix >= 9_999:
        print_error("Too many conflicting output filenames.", ["Clean output folder or rename source data"])
        raise ControlledExit
    return unique_output_path(folder, name, suffix + 1)


def extract_pdf_pages(reader: PdfReader, start: int, end: int, output: Path):
    writer = PdfWriter()
    list(map(writer.add_page, reader.pages[start - 1:end]))
    with open(output, "wb") as f:
        writer.write(f)


# ---------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------
def split_and_rename_pdf():
    print(f"{INFO}Starting...{RST}")

    excel_was_created = ensure_excel_exists(EXCEL_FILENAME)
    pdf_path = check_pdf_files(BASE_DIR)

    if excel_was_created:
        return

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

    def process_row(item):
        idx, row = item
        pdf_start, pdf_end = int(row.pdf_start), int(row.pdf_end)

        if pdf_start < 1 or pdf_end > total_pages or pdf_start > pdf_end:
            print_error(f"Invalid page range in row {idx}.", ["Check pdf_start and pdf_end values"])
            raise ControlledExit

        output_path = output_folder / f"{row.output_name}.pdf"

        if output_path.exists() and not overwrite_all:
            output_path = unique_output_path(output_folder, row.output_name)

        extract_pdf_pages(reader, pdf_start, pdf_end, output_path)

        progress = idx / total
        bar = "█" * int(30 * progress) + "-" * (30 - int(30 * progress))
        print(f"\r{INFO}|{bar}| {idx}/{total}{RST}", end="")

    list(map(process_row, enumerate(df.itertuples(), start=1)))

    print(f"\n{OK}PDFs successfully created.{RST}")


# ---------------------------------------------------------------------
if __name__ == "__main__":
    try:
        if not bootstrap_environment():
            raise ControlledExit
        split_and_rename_pdf()
    except ControlledExit:
        pass
    except Exception as e:
        print_error(
            f"Unexpected error: {e}",
            ["Ensure the PDF is closed", "Check write permissions"]
        )
