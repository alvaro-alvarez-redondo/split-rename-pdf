from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path
from typing import Final

# ---------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------
BASE_DIR: Final[Path] = Path(__file__).parent.resolve()
EXCEL_FILENAME: Final[Path] = BASE_DIR / "split-rename-pdf-mapping.xlsx"
OUTPUT_PATTERN: Final[str] = "{yearbook}_{category}_{year}_{first_page}_{last_page}_{product}"
REQUIRED_COLUMNS: Final[list[str]] = [
    "yearbook",
    "year",
    "category",
    "products",
    "yearbook_start",
    "yearbook_end",
    "pdf_start",
    "pdf_end",
]
FIELDS_TO_SANITIZE: Final[list[str]] = ["yearbook", "year", "category", "products"]
REQUIRED_PACKAGES: Final[list[str]] = ["pandas", "PyPDF2", "openpyxl"]
REQUIREMENTS_FILE: Final[Path] = BASE_DIR / "requirements.txt"
MAX_FILENAME_SUFFIX: Final[int] = 9_999

# ---------------------------------------------------------------------
# ANSI colors
# ---------------------------------------------------------------------
ERR: Final[str] = "\033[1;31m"
WARN: Final[str] = "\033[1;33m"
INFO: Final[str] = "\033[1;34m"
HELP: Final[str] = "\033[1;37m"
OK: Final[str] = "\033[1;32m"
RST: Final[str] = "\033[0m"


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
        print(f"{ERR}Error: Python 3.8+ is required{RST}")
        return False

    if not REQUIREMENTS_FILE.exists():
        REQUIREMENTS_FILE.write_text("\n".join(REQUIRED_PACKAGES), encoding="utf-8")
        print(
            f"{WARN}'requirements.txt' created. "
            "The script will attempt to install missing packages automatically."
            f"{RST}"
        )

    missing_modules = [package for package in REQUIRED_PACKAGES if not is_module_available(package)]

    if missing_modules:
        print(f"{WARN}Installing missing packages: {', '.join(missing_modules)}...{RST}")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", *missing_modules])
        except subprocess.CalledProcessError:
            print(f"{ERR}Failed to install packages: {', '.join(missing_modules)}{RST}")
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
def print_error(message: str, help_lines: list[str]) -> None:
    formatted_help = "\n".join(f"{HELP}• {line}{RST}" for line in help_lines)
    print(f"{ERR}{message}{RST}\n{formatted_help}")


def sanitize_filename(value: str) -> str:
    cleaned = re.sub(r'[\/:*?"<>|\s]+', "_", value.strip())
    return cleaned.strip("_").lower()


def ask_yes_no(prompt: str) -> bool:
    choice = input(prompt).strip().lower()
    if choice not in {"y", "n"}:
        print_error("Invalid input.", ["Please enter only 'y' or 'n'"])
        raise ControlledExit
    return choice == "y"


def check_pdf_files(base_dir: Path) -> Path:
    pdfs = list(base_dir.glob("*.pdf"))
    if len(pdfs) != 1:
        print_error("Expected exactly one PDF file.", ["Place only one .pdf file in the script directory"])
        raise ControlledExit
    return pdfs[0]


def create_output_folder(base_dir: Path, folder_name: str) -> Path:
    folder = base_dir / folder_name
    folder.mkdir(exist_ok=True)
    print(f"{INFO}Output folder '{folder.name}' ready{RST}")
    return folder


def resolve_output_folder_name(df: pd.DataFrame) -> str:
    folder_keys = (
        df[["yearbook", "year"]]
        .dropna(subset=["yearbook", "year"])
        .astype("string")
        .apply(lambda col: col.str.strip())
        .drop_duplicates()
    )

    if folder_keys.empty:
        print_error(
            "Unable to determine output folder name from Excel.",
            ["Ensure 'yearbook' and 'year' contain values"],
        )
        raise ControlledExit

    if len(folder_keys) > 1:
        print_error(
            "Multiple yearbook/year combinations found.",
            ["Use a single yearbook and year combination per run"],
        )
        raise ControlledExit

    record = folder_keys.iloc[0]
    yearbook = sanitize_filename(str(record["yearbook"]))
    year = sanitize_filename(str(record["year"]))
    return f"{yearbook}_extracted_pages_{year}"


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
    candidates = [path_item for path_item in target_path.parent.glob("*.xlsx") if path_item != target_path]
    valid_files = [candidate for candidate in candidates if is_valid_mapping_excel(candidate)]

    if len(valid_files) == 1:
        valid_files[0].rename(target_path)
        print(f"{INFO}Excel file found and renamed as '{target_path.name}'{RST}")
        return True

    if len(valid_files) > 1:
        print_error(
            "Multiple valid Excel files found.",
            ["Leave only one Excel with the required columns", f"Expected name: {target_path.name}"],
        )
        raise ControlledExit

    return False


def ensure_excel_exists(path: Path) -> bool:
    """Ensure mapping file exists. Returns True if created during this run."""
    if path.exists():
        return False

    if find_and_rename_valid_excel(path):
        return False

    pd.DataFrame(columns=REQUIRED_COLUMNS).to_excel(path, index=False)
    print_error(f"Excel file '{path.name}' was created.", ["Fill it with data and run the script again"])
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
        },
    )

    missing = set(REQUIRED_COLUMNS) - set(df.columns)
    if missing or df.empty:
        print_error("Excel validation failed.", ["Verify required columns exist", "Ensure the file is not empty"])
        raise ControlledExit

    return df


# ---------------------------------------------------------------------
# Data processing
# ---------------------------------------------------------------------
def handle_empty_products(df: pd.DataFrame) -> pd.DataFrame:
    empty = df["products"].isna() | (df["products"].astype(str).str.strip() == "")
    if not empty.any():
        return df

    intentional = ask_yes_no(f"{WARN}{empty.sum()} empty 'products'. Is this intentional? (y/n): {RST}")

    if not intentional:
        print_error("Empty products detected.", ["Fill all 'products' cells in Excel and rerun"])
        raise ControlledExit

    df.loc[empty, "products"] = ""
    return df


def generate_output_name(row) -> str:
    data = {field: sanitize_filename(str(getattr(row, field))) for field in FIELDS_TO_SANITIZE}

    if data["products"] == "":
        return (
            f"{data['yearbook']}_{data['category']}_{data['year']}_"
            f"{int(row.yearbook_start)}_{int(row.yearbook_end)}"
        )

    return OUTPUT_PATTERN.format(
        yearbook=data["yearbook"],
        category=data["category"],
        year=data["year"],
        first_page=int(row.yearbook_start),
        last_page=int(row.yearbook_end),
        product=data["products"],
    )


def unique_output_path(folder: Path, name: str) -> Path:
    for suffix in range(MAX_FILENAME_SUFFIX + 1):
        candidate = folder / (f"{name}.pdf" if suffix == 0 else f"{name}_{suffix}.pdf")
        if not candidate.exists():
            return candidate

    print_error("Too many conflicting output filenames.", ["Clean output folder or rename source data"])
    raise ControlledExit


def extract_pdf_pages(reader: PdfReader, start: int, end: int, output: Path) -> None:
    writer = PdfWriter()
    for page in reader.pages[start - 1 : end]:
        writer.add_page(page)

    with output.open("wb") as file_handle:
        writer.write(file_handle)


# ---------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------
def split_and_rename_pdf() -> None:
    print(f"{INFO}Starting...{RST}")

    excel_was_created = ensure_excel_exists(EXCEL_FILENAME)
    pdf_path = check_pdf_files(BASE_DIR)

    if excel_was_created:
        return

    df = load_excel(EXCEL_FILENAME)
    df = handle_empty_products(df)
    output_folder_name = resolve_output_folder_name(df)
    output_folder = create_output_folder(BASE_DIR, output_folder_name)

    reader = PdfReader(pdf_path)
    total_pages = len(reader.pages)

    df["output_name"] = df.apply(generate_output_name, axis=1)

    existing = df["output_name"].map(lambda name: (output_folder / f"{name}.pdf").exists())

    overwrite_all = False
    if existing.any():
        overwrite_all = ask_yes_no(f"{WARN}{existing.sum()} files already exist. Overwrite all? (y/n): {RST}")

    total = len(df)

    for idx, row in enumerate(df.itertuples(), start=1):
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

    print(f"\n{OK}PDFs successfully created.{RST}")


# ---------------------------------------------------------------------
if __name__ == "__main__":
    try:
        if not bootstrap_environment():
            raise ControlledExit
        split_and_rename_pdf()
    except ControlledExit:
        pass
    except Exception as exc:
        print_error(f"Unexpected error: {exc}", ["Ensure the PDF is closed", "Check write permissions"])
