from __future__ import annotations

import argparse
import importlib
import logging
import re
import subprocess
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any

# ---------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------
BASE_DIR = Path(__file__).parent.resolve()
EXCEL_FILENAME = BASE_DIR / "split-rename-pdf-mapping.xlsx"
OUTPUT_PATTERN = "{yearbook}_{category}_{year}_{first_page}_{last_page}_{product}"
REQUIRED_COLUMNS = [
    "yearbook",
    "year",
    "category",
    "products",
    "yearbook_start",
    "yearbook_end",
    "pdf_start",
    "pdf_end",
]
FIELDS_TO_SANITIZE = ["yearbook", "year", "category", "products"]
REQUIRED_PACKAGES = ["pandas", "PyPDF2", "openpyxl"]
MAX_FILENAME_LENGTH = 180
MAX_COLLISION_ATTEMPTS = 9_999

LOGGER = logging.getLogger("split_rename_pdf")


class UserFacingError(Exception):
    """Raised for validation and expected user-facing errors."""


@dataclass(frozen=True)
class RuntimeDependencies:
    """Container for runtime-only dependencies."""

    pandas: Any
    pdf_reader: Any
    pdf_writer: Any


@dataclass(frozen=True)
class CliOptions:
    """Command-line options for the split/rename workflow."""

    pdf: Path | None
    mapping: Path | None
    output: Path | None
    overwrite: bool
    dry_run: bool
    verbose: bool


def configure_logging(verbose: bool) -> None:
    """Configure global logging output for CLI usage."""
    level = logging.INFO if verbose else logging.WARNING
    logging.basicConfig(level=level, format="%(levelname)s: %(message)s")


def parse_cli_args(argv: list[str] | None = None) -> CliOptions:
    """Parse CLI arguments and return normalized options."""
    parser = argparse.ArgumentParser(
        description=(
            "Split one PDF into multiple files based on page ranges defined in an Excel mapping file."
        ),
        epilog=(
            "Examples:\n"
            "  python split-rename-pdf.py\n"
            "  python split-rename-pdf.py --pdf ./book.pdf --mapping ./mapping.xlsx\n"
            "  python split-rename-pdf.py --output ./out --overwrite --verbose\n"
            "  python split-rename-pdf.py --dry-run"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    parser.add_argument(
        "--pdf",
        type=Path,
        default=None,
        help="Path to source PDF. Default: auto-detect exactly one .pdf in the script folder.",
    )
    parser.add_argument(
        "--mapping",
        type=Path,
        default=None,
        help="Path to Excel mapping file. Default: split-rename-pdf-mapping.xlsx in script folder.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Output directory. Default: a folder named after the input PDF in script folder.",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite existing output files without prompting.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Validate mapping and print planned output files without writing anything.",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Enable informative logs (INFO, WARNING, ERROR).",
    )

    args = parser.parse_args(argv)
    return CliOptions(
        pdf=args.pdf.resolve() if args.pdf else None,
        mapping=args.mapping.resolve() if args.mapping else None,
        output=args.output.resolve() if args.output else None,
        overwrite=args.overwrite,
        dry_run=args.dry_run,
        verbose=args.verbose,
    )


def is_module_available(module_name: str) -> bool:
    """Return ``True`` when the module can be imported."""
    try:
        importlib.import_module(module_name)
    except ModuleNotFoundError:
        return False
    return True


def bootstrap_environment() -> RuntimeDependencies:
    """Prepare runtime dependencies."""
    if sys.version_info < (3, 8):
        raise UserFacingError("Python 3.8+ is required. Use a newer Python interpreter and rerun.")

    missing_modules = [pkg for pkg in REQUIRED_PACKAGES if not is_module_available(pkg)]
    if missing_modules:
        LOGGER.warning("Installing missing packages: %s", ", ".join(missing_modules))
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", *missing_modules])
        except subprocess.CalledProcessError as exc:
            raise UserFacingError(
                f"Failed to install required packages: {', '.join(missing_modules)}"
            ) from exc

    import pandas as pandas_module
    from PyPDF2 import PdfReader as reader_module, PdfWriter as writer_module

    return RuntimeDependencies(
        pandas=pandas_module,
        pdf_reader=reader_module,
        pdf_writer=writer_module,
    )


def ask_yes_no(prompt: str) -> bool:
    """Prompt the user for an explicit ``y`` or ``n`` response."""
    choice = input(prompt).strip().lower()
    if choice not in {"y", "n"}:
        raise UserFacingError("Invalid input. Please enter only 'y' or 'n'.")
    return choice == "y"


def resolve_pdf_path(base_dir: Path, pdf_arg: Path | None) -> Path:
    """Resolve source PDF from CLI argument or folder auto-detection."""
    if pdf_arg is not None:
        if not pdf_arg.exists() or not pdf_arg.is_file():
            raise UserFacingError(f"Provided --pdf path does not exist: {pdf_arg}")
        if pdf_arg.suffix.lower() != ".pdf":
            raise UserFacingError(f"Provided --pdf path is not a PDF file: {pdf_arg}")
        return pdf_arg

    pdfs = list(base_dir.glob("*.pdf"))
    if len(pdfs) != 1:
        raise UserFacingError(
            "Expected exactly one PDF file in the script directory. "
            "Place one .pdf file there or pass --pdf PATH."
        )
    return pdfs[0]


def resolve_mapping_path(mapping_arg: Path | None) -> Path:
    """Resolve mapping path and enforce default behavior compatibility."""
    return mapping_arg if mapping_arg is not None else EXCEL_FILENAME


def create_output_folder(base_dir: Path, pdf_path: Path, output_arg: Path | None, dry_run: bool) -> Path:
    """Resolve output folder and create it unless dry-run mode is active."""
    folder = output_arg if output_arg is not None else base_dir / pdf_path.stem
    if dry_run:
        LOGGER.info("Dry-run: output folder would be %s", folder)
        return folder

    try:
        folder.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        raise UserFacingError(f"Unable to create output folder: {folder}") from exc

    LOGGER.info("Output folder ready: %s", folder)
    return folder


def is_valid_mapping_excel(excel_path: Path, pandas_module: Any) -> bool:
    """Check whether an Excel file has all required mapping columns."""
    try:
        columns = pandas_module.read_excel(excel_path, nrows=0).columns
    except Exception:
        return False
    return set(REQUIRED_COLUMNS).issubset(columns)


def find_and_rename_valid_excel(target_path: Path, pandas_module: Any, dry_run: bool) -> bool:
    """Find one valid workbook and rename it to expected mapping filename."""
    candidates = [
        path_item for path_item in target_path.parent.glob("*.xlsx") if path_item != target_path
    ]
    valid_files = [path for path in candidates if is_valid_mapping_excel(path, pandas_module)]

    if len(valid_files) == 1:
        if dry_run:
            LOGGER.info(
                "Dry-run: would rename %s to %s", valid_files[0].name, target_path.name
            )
            return True
        valid_files[0].rename(target_path)
        LOGGER.info("Excel file found and renamed as '%s'", target_path.name)
        return True

    if len(valid_files) > 1:
        raise UserFacingError(
            "Multiple valid Excel files found. Leave only one Excel with required columns "
            f"or use --mapping. Expected default name: {target_path.name}"
        )

    return False


def ensure_mapping_exists(path: Path, pandas_module: Any, allow_auto_rename: bool, dry_run: bool) -> None:
    """Ensure mapping file exists while respecting dry-run no-write behavior."""
    if path.exists():
        return

    if allow_auto_rename and find_and_rename_valid_excel(path, pandas_module, dry_run=dry_run):
        return

    if dry_run:
        raise UserFacingError(
            f"Mapping file not found: {path}. In dry-run mode no files are created; provide --mapping."
        )

    pandas_module.DataFrame(columns=REQUIRED_COLUMNS).to_excel(path, index=False)
    raise UserFacingError(
        f"Excel mapping file '{path.name}' was created. Fill it with data and run again."
    )


def load_excel(path: Path, pandas_module: Any) -> Any:
    """Load mapping workbook with early schema and value validation."""
    try:
        dataframe = pandas_module.read_excel(path)
    except Exception as exc:
        raise UserFacingError(f"Failed to read Excel mapping file: {path}") from exc

    missing_columns = set(REQUIRED_COLUMNS) - set(dataframe.columns)
    if missing_columns:
        missing = ", ".join(sorted(missing_columns))
        raise UserFacingError(f"Excel mapping is missing required columns: {missing}")

    if dataframe.empty:
        raise UserFacingError("Excel mapping file is empty. Add at least one mapping row.")

    # Normalize text columns.
    for text_column in ("yearbook", "year", "category", "products"):
        dataframe[text_column] = dataframe[text_column].astype("string")

    # Coerce numeric columns and validate all values are numeric and present.
    numeric_columns = ["yearbook_start", "yearbook_end", "pdf_start", "pdf_end"]
    for column in numeric_columns:
        dataframe[column] = pandas_module.to_numeric(dataframe[column], errors="coerce")

    invalid_numeric = dataframe[numeric_columns].isna().any(axis=1)
    if invalid_numeric.any():
        invalid_rows = ", ".join(str(index + 2) for index in dataframe.index[invalid_numeric])
        raise UserFacingError(
            "Non-numeric or empty page values found in rows: "
            f"{invalid_rows}. Check yearbook_start/yearbook_end/pdf_start/pdf_end."
        )

    for column in numeric_columns:
        dataframe[column] = dataframe[column].astype("int64")

    invalid_yearbook_range = (
        (dataframe["yearbook_start"] < 1) | (dataframe["yearbook_end"] < dataframe["yearbook_start"])
    )
    if invalid_yearbook_range.any():
        invalid_rows = ", ".join(str(index + 2) for index in dataframe.index[invalid_yearbook_range])
        raise UserFacingError(
            f"Invalid yearbook page range in rows: {invalid_rows}. "
            "Ensure yearbook_start >= 1 and yearbook_end >= yearbook_start."
        )

    invalid_pdf_range = (
        (dataframe["pdf_start"] < 1) | (dataframe["pdf_end"] < dataframe["pdf_start"])
    )
    if invalid_pdf_range.any():
        invalid_rows = ", ".join(str(index + 2) for index in dataframe.index[invalid_pdf_range])
        raise UserFacingError(
            f"Invalid PDF page range in rows: {invalid_rows}. "
            "Ensure pdf_start >= 1 and pdf_end >= pdf_start."
        )

    return dataframe


def handle_empty_products(dataframe: Any) -> Any:
    """Validate empty product rows and ask user for explicit confirmation."""
    empty_products = dataframe["products"].isna() | (dataframe["products"].astype(str).str.strip() == "")
    if not empty_products.any():
        return dataframe

    intentional = ask_yes_no(
        f"WARNING: {empty_products.sum()} empty 'products' values found. Is this intentional? (y/n): "
    )
    if not intentional:
        raise UserFacingError("Empty products detected. Fill all 'products' cells and rerun.")

    dataframe.loc[empty_products, "products"] = ""
    return dataframe


def sanitize_filename(value: str) -> str:
    """Sanitize a path component for safer cross-platform filenames."""
    cleaned = re.sub(r"[\x00-\x1f\x7f]+", "", value.strip())
    cleaned = re.sub(r"[\\/:*?\"<>|]+", "_", cleaned)
    cleaned = re.sub(r"\s+", "_", cleaned)
    cleaned = cleaned.strip("._ ")
    cleaned = cleaned.lower()

    if not cleaned:
        cleaned = "untitled"

    if len(cleaned) > MAX_FILENAME_LENGTH:
        cleaned = cleaned[:MAX_FILENAME_LENGTH].rstrip("._ ")

    return cleaned or "untitled"


def generate_output_name(row: Any) -> str:
    """Create deterministic output filename (without extension) for one row."""
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
    """Return a non-conflicting output path using iterative suffixes."""
    for suffix in range(0, MAX_COLLISION_ATTEMPTS + 1):
        candidate = folder / (f"{name}.pdf" if suffix == 0 else f"{name}_{suffix}.pdf")
        if not candidate.exists():
            return candidate

    raise UserFacingError(
        "Too many conflicting output filenames. Clean output folder or rename source data."
    )


def extract_pdf_pages(reader: Any, writer_cls: Any, start: int, end: int, output: Path) -> None:
    """Extract inclusive page range [start, end] and write to ``output``."""
    writer = writer_cls()
    for page in reader.pages[start - 1 : end]:
        writer.add_page(page)

    with output.open("wb") as file_obj:
        writer.write(file_obj)


def find_overlapping_ranges(dataframe: Any) -> list[tuple[int, int, int, int, int, int]]:
    """Return overlapping PDF page ranges as row-pair tuples.

    Each tuple contains:
    ``(row_a, start_a, end_a, row_b, start_b, end_b)``.
    """
    ranges: list[tuple[int, int, int]] = []
    for row_number, row in enumerate(dataframe.itertuples(index=False), start=2):
        ranges.append((row_number, int(row.pdf_start), int(row.pdf_end)))

    overlaps: list[tuple[int, int, int, int, int, int]] = []
    for index, (row_a, start_a, end_a) in enumerate(ranges):
        for row_b, start_b, end_b in ranges[index + 1 :]:
            if start_a <= end_b and start_b <= end_a:
                overlaps.append((row_a, start_a, end_a, row_b, start_b, end_b))

    return overlaps


def render_progress_line(
    rows_done: int,
    rows_total: int,
    pages_done: int,
    pages_total: int,
    bar_width: int = 20,
) -> str:
    """Build a text progress line with row and page progress."""
    fraction = rows_done / rows_total if rows_total else 1.0
    filled = int(bar_width * fraction)
    bar = "█" * filled + "-" * (bar_width - filled)
    return f"[{bar}] {rows_done}/{rows_total} rows | Pages: {pages_done}/{pages_total}"


def log_final_summary(
    pdf_path: Path,
    rows_processed: int,
    output_files: int,
    extracted_pages: int,
    elapsed_seconds: float,
) -> None:
    """Log a final operation summary in a user-friendly format."""
    LOGGER.warning("Summary:\n---------")
    LOGGER.warning("PDF source: %s", pdf_path.name)
    LOGGER.warning("Rows processed: %s", rows_processed)
    LOGGER.warning("Output files: %s", output_files)
    LOGGER.warning("Total pages extracted: %s", extracted_pages)
    LOGGER.warning("Time elapsed: %.2fs", elapsed_seconds)


def split_and_rename_pdf(dependencies: RuntimeDependencies, options: CliOptions) -> None:
    """Orchestrate workbook validation, page splitting, and file generation."""
    start_time = time.perf_counter()

    mapping_path = resolve_mapping_path(options.mapping)
    pdf_path = resolve_pdf_path(BASE_DIR, options.pdf)
    output_folder = create_output_folder(BASE_DIR, pdf_path, options.output, options.dry_run)

    LOGGER.info("Starting processing")
    LOGGER.info("PDF source: %s", pdf_path)
    LOGGER.info("Mapping file: %s", mapping_path)

    ensure_mapping_exists(
        mapping_path,
        dependencies.pandas,
        allow_auto_rename=options.mapping is None,
        dry_run=options.dry_run,
    )

    dataframe = load_excel(mapping_path, dependencies.pandas)
    dataframe = handle_empty_products(dataframe)

    overlaps = find_overlapping_ranges(dataframe)
    if overlaps:
        LOGGER.warning("Overlapping page ranges detected:")
        for row_a, start_a, end_a, row_b, start_b, end_b in overlaps:
            LOGGER.warning(
                "Row %s (pages %s-%s) overlaps with Row %s (pages %s-%s)",
                row_a,
                start_a,
                end_a,
                row_b,
                start_b,
                end_b,
            )

        continue_with_overlaps = ask_yes_no(
            "WARNING: overlapping ranges found. Continue processing anyway? (y/n): "
        )
        if not continue_with_overlaps:
            raise UserFacingError("Execution stopped due to overlapping page ranges.")


    output_names: list[str] = []
    for row in dataframe.itertuples(index=False):
        output_names.append(generate_output_name(row))
    dataframe["output_name"] = output_names

    existing_count = 0
    for name in output_names:
        if (output_folder / f"{name}.pdf").exists():
            existing_count += 1

    overwrite_all = options.overwrite
    if existing_count and not overwrite_all and not options.dry_run:
        overwrite_all = ask_yes_no(
            f"WARNING: {existing_count} files already exist. Overwrite all? (y/n): "
        )

    planned_writes = 0
    extracted_pages = 0
    reader = dependencies.pdf_reader(pdf_path)
    total_pages = len(reader.pages)
    rows_total = len(dataframe)
    pages_total = 0
    for row in dataframe.itertuples(index=False):
        pages_total += int(row.pdf_end) - int(row.pdf_start) + 1

    for row_number, row in enumerate(dataframe.itertuples(index=False), start=2):
        pdf_start = int(row.pdf_start)
        pdf_end = int(row.pdf_end)
        pages_in_row = pdf_end - pdf_start + 1

        if pdf_end > total_pages:
            raise UserFacingError(
                f"Row {row_number} has invalid PDF range {pdf_start}-{pdf_end}. "
                f"Input PDF has only {total_pages} pages."
            )

        output_path = output_folder / f"{row.output_name}.pdf"
        if output_path.exists() and not overwrite_all:
            output_path = unique_output_path(output_folder, row.output_name)

        if options.dry_run:
            LOGGER.warning(
                "Dry-run: row=%s pages=%s-%s output=%s", row_number, pdf_start, pdf_end, output_path
            )
            planned_writes += 1
            extracted_pages += pages_in_row
            continue

        extract_pdf_pages(reader, dependencies.pdf_writer, pdf_start, pdf_end, output_path)
        planned_writes += 1
        extracted_pages += pages_in_row
        LOGGER.info("Created: %s", output_path)

        if not options.verbose:
            progress_line = render_progress_line(
                rows_done=planned_writes,
                rows_total=rows_total,
                pages_done=extracted_pages,
                pages_total=pages_total,
            )
            sys.stdout.write(f"\r{progress_line}")
            sys.stdout.flush()

    if not options.dry_run and not options.verbose and rows_total:
        sys.stdout.write("\n")

    if options.dry_run:
        LOGGER.warning("Dry-run completed. Planned %s output files. No files were written.", planned_writes)
    else:
        LOGGER.warning("Completed successfully. Created %s PDF files.", planned_writes)

    elapsed_seconds = time.perf_counter() - start_time
    rows_processed = len(dataframe)
    log_final_summary(
        pdf_path=pdf_path,
        rows_processed=rows_processed,
        output_files=planned_writes,
        extracted_pages=extracted_pages,
        elapsed_seconds=elapsed_seconds,
    )


def main(argv: list[str] | None = None) -> None:
    """CLI entrypoint."""
    options = parse_cli_args(argv)
    configure_logging(options.verbose)

    try:
        dependencies = bootstrap_environment()
        split_and_rename_pdf(dependencies, options)
    except UserFacingError as exc:
        LOGGER.error("%s", exc)
        sys.exit(1)
    except Exception as exc:  # noqa: BLE001
        LOGGER.error("Unexpected internal error: %s", exc)
        if options.verbose:
            LOGGER.exception("Traceback:")
        sys.exit(1)


if __name__ == "__main__":
    main()
