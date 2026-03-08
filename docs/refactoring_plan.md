# Refactoring Plan: `split-rename-pdf`

## Goals

- Improve maintainability by separating concerns.
- Improve readability with small, focused modules and explicit naming.
- Improve modularity and testability by isolating pure logic from side effects.
- Improve error handling and user-facing feedback.
- Reduce global mutable state.
- Preserve current CLI behavior and workflow.

---

## 1) Proposed folder structure

```text
split-rename-pdf/
├── pyproject.toml
├── README.md
├── requirements.txt
├── src/
│   └── split_rename_pdf/
│       ├── __init__.py
│       ├── __main__.py
│       ├── cli.py
│       ├── config.py
│       ├── models.py
│       ├── errors.py
│       ├── bootstrap.py
│       ├── excel_mapping.py
│       ├── pdf_ops.py
│       ├── naming.py
│       └── workflow.py
├── tests/
│   ├── test_naming.py
│   ├── test_output_paths.py
│   ├── test_excel_validation.py
│   └── test_workflow_smoke.py
└── split-rename-pdf.py  # optional compatibility shim (phase-out path)
```

Notes:
- Keep packaging minimal via `src/` layout.
- Preserve lightweight usage by supporting `python split-rename-pdf.py` during transition.

---

## 2) Suggested modules and responsibilities

### `config.py`
- Constants and immutable config defaults:
  - `BASE_DIR`, `EXCEL_FILENAME`, `OUTPUT_PATTERN`, `REQUIRED_COLUMNS`, `REQUIRED_PACKAGES`.
- Optional dataclass-based runtime settings:
  - `AppConfig` for resolved paths and patterns.

### `errors.py`
- Domain exceptions:
  - `ControlledExit`, `ValidationError`, `DependencyError`.
- Keeps error taxonomy explicit and testable.

### `bootstrap.py`
- Runtime environment preparation:
  - Python version check.
  - Dependency detection/installation.
- Returns typed runtime handles (e.g., pandas/PyPDF2 references) without module-level mutation.

### `excel_mapping.py`
- Excel-related functions:
  - Validate headers.
  - Create template workbook if missing.
  - Load mapping dataframe with explicit dtypes.
  - Validate non-empty and column presence.

### `pdf_ops.py`
- PDF-related operations:
  - Locate source PDF.
  - Split page ranges.
  - Write output files.

### `naming.py`
- Filename sanitization and output-name generation.
- Unique output path resolution.

### `workflow.py`
- Orchestration service:
  - Combine mapping + PDF operations + prompts.
  - Minimal business flow logic.

### `cli.py`
- User I/O boundary:
  - Argument parsing (optional future flags).
  - Prompting (`y/n`) and progress output.
  - Top-level exception mapping to friendly messages.

### `__main__.py`
- Entrypoint to run CLI package with:
  - `python -m split_rename_pdf`

---

## 3) Recommended classes and functions

Keep class usage minimal to avoid over-engineering.

### Classes
- `RuntimeDependencies` (dataclass): runtime-imported modules/classes.
- `AppConfig` (dataclass): resolved paths and naming pattern.

### Core functions (examples)
- `bootstrap_environment(config: AppConfig) -> RuntimeDependencies`
- `ensure_mapping_excel_exists(excel_path: Path, required_columns: list[str], pandas_module: Any) -> bool`
- `load_mapping_dataframe(excel_path: Path, pandas_module: Any) -> Any`
- `validate_page_range(pdf_start: int, pdf_end: int, total_pages: int) -> None`
- `extract_pdf_pages(reader: Any, writer_class: Any, page_start: int, page_end: int, output_path: Path) -> None`
- `generate_output_name(row: Any, output_pattern: str) -> str`
- `build_unique_output_path(output_dir: Path, base_name: str, max_suffix: int = 9999) -> Path`
- `run_split_workflow(config: AppConfig, dependencies: RuntimeDependencies) -> None`

---

## 4) Configuration handling strategy

Use a two-tier configuration model:

1. **Static defaults as constants** (UPPER_CASE in `config.py`) for predictable behavior.
2. **Resolved runtime config dataclass (`AppConfig`)** passed explicitly to functions.

Benefits:
- Eliminates hidden coupling.
- Supports future CLI flags without invasive rewrites.
- Keeps behavior deterministic and test-friendly.

Example:
- Resolve `base_dir` once at startup.
- Derive `excel_filename` and output folder from `base_dir`.
- Pass config into workflow rather than reading globals directly.

---

## 5) CLI architecture recommendation

### Keep the CLI simple
- Continue default no-argument workflow for backward compatibility.
- Optional non-breaking flags can be added later:
  - `--base-dir`
  - `--mapping-file`
  - `--yes` (non-interactive overwrite/empty-products confirmation)

### Execution model
1. Parse args.
2. Build `AppConfig`.
3. Bootstrap dependencies.
4. Run workflow.
5. Catch known exceptions and print concise, actionable messages.

### Error handling
- Convert expected failures to `ControlledExit` with user guidance.
- Preserve traceback only for unexpected errors (or behind a debug flag).

---

## 6) Example refactored function signatures (snake_case)

```python
def bootstrap_environment(config: AppConfig) -> RuntimeDependencies:
    ...


def ensure_mapping_excel_exists(
    excel_path: Path,
    required_columns: list[str],
    pandas_module: Any,
) -> bool:
    ...


def generate_output_name(row: Any, output_pattern: str) -> str:
    ...


def build_unique_output_path(
    output_dir: Path,
    base_name: str,
    max_suffix: int = 9_999,
) -> Path:
    ...


def run_split_workflow(config: AppConfig, dependencies: RuntimeDependencies) -> None:
    ...
```

All names follow strict requirements:
- Functions/variables/parameters/modules: `snake_case`
- Classes: `PascalCase`
- Constants: `UPPER_CASE`

---

## 7) Specific renaming suggestions (PEP 8 alignment)

Most names already follow PEP 8. The following targeted renames improve clarity and consistency further:

- `split_and_rename_pdf` → `run_split_workflow`
  - Clarifies orchestration role.
- `check_pdf_files` → `find_single_pdf_file`
  - Indicates both lookup and cardinality expectation.
- `load_excel` → `load_mapping_dataframe`
  - Makes return type/intent explicit.
- `ensure_excel_exists` → `ensure_mapping_excel_exists`
  - Disambiguates purpose.
- `find_and_rename_valid_excel` → `find_and_prepare_mapping_excel`
  - Better reflects behavior.
- `unique_output_path` → `build_unique_output_path`
  - Reads like a pure path-construction function.
- Parameter `deps` → `dependencies`
  - Improves readability in call sites.
- Local variable `dataframe` can remain `dataframe`; avoid abbreviated names like `df` in orchestration code.

No class/constants renaming needed if these remain:
- `RuntimeDependencies` (PascalCase)
- `BASE_DIR`, `EXCEL_FILENAME`, `OUTPUT_PATTERN`, `REQUIRED_COLUMNS` (UPPER_CASE)

---

## Lightweight implementation sequence (recommended)

1. **Introduce package layout** under `src/split_rename_pdf/` while keeping current script as compatibility shim.
2. **Move pure helpers first** (`naming.py`, simple validators), add unit tests.
3. **Move Excel/PDF adapters** (`excel_mapping.py`, `pdf_ops.py`) with unchanged behavior.
4. **Extract workflow orchestration** into `workflow.py`.
5. **Add CLI wrapper** in `cli.py` + `__main__.py`.
6. **Keep backward-compatible script** calling package CLI until users migrate.

This sequence minimizes risk while preserving behavior and user workflow.
