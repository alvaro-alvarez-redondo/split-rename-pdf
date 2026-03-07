# PDF Split & Rename CLI

A small Python command-line utility for splitting one source PDF into many output PDFs using row-by-row instructions from an Excel mapping file.

It is designed for repeatable office workflows where naming consistency and page-range accuracy are important (for example: catalogs, yearbooks, or reports).

## Features

- Splits exactly one input PDF into multiple PDFs.
- Uses an Excel file (`rename-pdf-mapping.xlsx`) as the split-and-rename mapping source.
- Auto-creates an Excel template when missing.
- Sanitizes output names (lowercase, safe characters, trimmed separators).
- Handles duplicate output names by prompting for overwrite or creating unique suffixes.
- Validates page ranges before extraction.
- Shows progress in the terminal.

## Requirements

- Python 3.8+
- Dependencies:
  - `pandas`
  - `PyPDF2`
  - `openpyxl`

> On first run, the script can generate `requirements.txt` and attempt to install missing packages.

## Installation

1. Clone or download this repository.
2. Place `split-rename-pdf.py` in your working folder.
3. Install dependencies (recommended):

```bash
pip install -r requirements.txt
```

If `requirements.txt` does not yet exist, run the script once to generate it.

## Quick Start

1. Put files in the same folder as the script:
   - `split-rename-pdf.py`
   - one input PDF (exactly one `.pdf`)
   - `rename-pdf-mapping.xlsx` (or let the script create it)
2. Run:

```bash
python split-rename-pdf.py
```

3. If the Excel mapping file is missing, the script creates a template and informs you.
4. Fill the Excel rows and run again.
5. The script creates an output directory named after the input PDF filename.

## Excel Mapping Format

The Excel file **must** be named:

- `rename-pdf-mapping.xlsx`

Required columns:

- `yearbook`
- `year`
- `category`
- `products`
- `yearbook_start`
- `yearbook_end`
- `pdf_start`
- `pdf_end`

### Column meanings

| Column | Purpose |
|---|---|
| `yearbook` | Main document identifier used in output filename |
| `year` | Year label used in output filename |
| `category` | Section/category label used in output filename |
| `products` | Product label used in output filename (can be blank if intentional) |
| `yearbook_start` | Real-world start page shown in filename |
| `yearbook_end` | Real-world end page shown in filename |
| `pdf_start` | Start page to extract from the source PDF |
| `pdf_end` | End page to extract from the source PDF |

### Example mapping rows

| yearbook | year | category | products | yearbook_start | yearbook_end | pdf_start | pdf_end |
|---|---|---|---|---:|---:|---:|---:|
| acme_catalog | 2024 | tools | hammer_pro | 10 | 15 | 1 | 6 |
| acme_catalog | 2024 | tools | drill_x | 16 | 20 | 7 | 11 |
| acme_catalog | 2024 | accessories |  | 21 | 24 | 12 | 15 |

## Output Structure

The script creates an output folder named after the input PDF stem:

```text
<source-pdf-name>/
```

Output filename pattern:

```text
{yearbook}_{category}_{year}_{first_page}_{last_page}_{product}.pdf
```

Example:

```text
acme_catalog_tools_2024_10_15_hammer_pro.pdf
```

If `products` is empty (and confirmed as intentional), the trailing product segment is omitted.

## Error Handling

The script checks and reports:

- Missing or invalid Excel structure.
- Empty mapping data.
- Invalid page ranges.
- Missing or multiple PDFs in the directory.
- Duplicate output filenames.

Messages are intentionally user-friendly and actionable.

## Troubleshooting

- **No Excel file appears?**
  - Run the script once in the working folder; it creates `rename-pdf-mapping.xlsx` when missing.
- **"Expected exactly one PDF file" error?**
  - Keep only one `.pdf` in the script directory.
- **Package import errors?**
  - Install dependencies manually:

```bash
pip install pandas PyPDF2 openpyxl
```

- **Page range errors?**
  - Confirm `pdf_start` and `pdf_end` are integers and inside the source PDF page count.

## License

No license file is currently included in this repository.
If you plan to distribute this project, add a license such as MIT, Apache-2.0, or GPL-3.0.
