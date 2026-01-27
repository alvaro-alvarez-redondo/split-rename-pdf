# PDF Split & Rename Tool

A Python utility that splits a single PDF into multiple PDFs based on page ranges defined in an Excel file, and renames the outputs using a standardized, sanitized naming convention.

This tool is designed for repeatable, automated workflows where large PDFs (such as catalogs, yearbooks, or reports) need to be segmented into well-defined documents.

---

## Features

- Splits **one input PDF** into multiple output PDFs
- Page ranges and metadata are defined via an **Excel mapping file**
- Automatic **file name sanitization**
- Optional handling of empty product names
- Detects existing output files and safely resolves name conflicts

---

## Project Structure

```text
project/
│
├── split_and_rename.py       # Main script
├── rename-pdf-mapping.xlsx   # Excel mapping file
├── input.pdf                 # Single source PDF (exactly one)
└── input/                    # Auto-created output folder
```

> ⚠️ Exactly **one PDF file** must exist in the script directory when running the tool.

---

## Requirements

- Dependencies:
  - `pandas`
  - `PyPDF2`
  - `openpyxl` (for Excel support)

---

## Excel Mapping File

The Excel file `rename-pdf-mapping.xlsx` controls how the PDF is split and renamed.

### Required Columns

| Column Name        | Description |
|-------------------|-------------|
| `yearbook`        | Main document identifier |
| `year`            | Year label |
| `category`        | Category or section name |
| `products`        | Product name (may be empty if intentional) |
| `yearbook_start`  | IRL start page (for naming only) |
| `yearbook_end`    | IRL end page (for naming only) |
| `pdf_start`       | Start page in the source PDF |
| `pdf_end`         | End page in the source PDF |

### Notes

- All numeric columns **must contain integers**
- Page ranges must be valid and within the total number of PDF pages
- Empty `products` values require explicit confirmation at runtime

---

## Output File Naming

Output filenames follow this pattern:

```text
{yearbook}_{category}_{year}_{first_page}_{last_page}_{product}.pdf
```

If `products` is empty, the product segment is omitted:

```text
yearbook_category_year_first_last.pdf
```

All components are automatically:
- Converted to lowercase
- Trimmed
- Sanitized to remove unsafe characters

---

## Usage

1. Place the script, Excel file, and **exactly one PDF** in the same directory
2. Fill in the Excel mapping file
3. Run the script
4. Output PDFs will be created in a folder named after the input PDF

---

## Runtime Behavior

- If the Excel file does not exist, a template is created and execution stops
- If output files already exist, you will be prompted to:
  - Overwrite all, or
  - Automatically generate unique filenames
- Any invalid configuration results in a clear error message and safe exit

---

## Error Handling

The script validates:

- Excel structure and required columns
- Empty or invalid data
- Page range correctness
- File system permissions
- Conflicting output filenames

All errors are reported with actionable guidance.

---

## Limitations

- Only **one input PDF** is supported per run
- PDF processing is sequential (due to library constraints)
- No GUI (CLI-only by design)

