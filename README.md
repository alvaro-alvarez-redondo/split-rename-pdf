# split-rename-pdf

A Python script to automatically split a single PDF into multiple files and rename them based on metadata provided in an Excel file. The script ensures safe filenames, validates page ranges, and provides progress feedback in the terminal.

---

## Features

- Split a single PDF into multiple PDFs based on page ranges.
- Rename output PDFs using a customizable naming pattern.
- Load metadata from an Excel file (`.xlsx`) with required columns.
- Validate Excel data and PDF page ranges.
- Handle existing files with overwrite or auto-increment options.
- Colored terminal output for errors, warnings, and info messages.
- Automatically creates an output folder.

---

## Extra information

### Configuration

1. Place exactly one PDF in the same folder as the script.
1. Prepare an Excel file named rename-pdf-mapping.xlsx in the same folder.

### Required Excel columns:
- "yearbook"
- "year"
- "category"
- "products"
- "yearbook_start"
- "yearbook_end"
- "pdf_start"
- "pdf_end"

### Optional: Customize output filename pattern in the script
```python
OUTPUT_PATTERN = "{yearbook}_{category}_{year}_{first_page}_{last_page}_{product}"
```

Fields to sanitize (replace invalid characters in filenames)
```python
FIELDS_TO_SANITIZE = ["yearboos", "year", "category", "products"]
```

### Steps performed:

1. Script checks the folder for exactly one PDF.
1. Loads the Excel mapping file.
1. Validates Excel columns and data.
1. Creates an output folder named after the input PDF.
1. Splits PDF pages according to the Excel file and renames them.

