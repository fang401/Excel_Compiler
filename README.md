# Excel Data Compilation Tool

A professional Streamlit application for uploading multiple CSV/Excel files,
reviewing their contents, and compiling them into a single Excel workbook —
one worksheet per uploaded file — **without altering any of the original
data**.

## Overview

Business users often receive several separate CSV/Excel exports (e.g. a
Non-Conformance report, a Portfolio Listing, an Ageing report) that need to
be combined into one workbook for review or distribution. This tool does
exactly that: it validates each uploaded file, shows a quick summary and
preview, and compiles everything into a single downloadable `.xlsx` file —
one tab per source file — with the source data preserved exactly as
uploaded.

## Features

- **Multiple file upload** — any combination of `.csv`, `.xlsx`, `.xls`
  files, any number, none required.
- **Validation** — empty files, unsupported formats, and corrupted files
  are caught and reported clearly, without crashing the app or silently
  skipping files.
- **Dataset summary dashboard** — automatically calculates business
  metrics for recognized file types:
  - *Non-Conformance Data*: Non-Compliance Projects, Status Report /
    Financials / Workplans Non-compliance counts.
  - *Portfolio Listing*: Total Projects.
  - Files that don't match a recognized type are still compiled — they
    simply don't get a metrics card.
- **Data preview** — pick any uploaded file and preview its first rows,
  row/column counts, and column names.
- **One-click compilation** — combines every successfully validated file
  into a single workbook, one worksheet per file, named after the original
  filename (extension removed, Excel's 31-character/invalid-character
  worksheet-name rules handled automatically).
- **Strict data integrity** — the app never modifies, cleans, sorts,
  filters, or reformats the uploaded data. What you upload is what ends up
  in the compiled workbook.

## Project Structure

```
excel-compilation-tool/
│
├── app.py                     # Streamlit application (UI + orchestration)
├── utils/
│   ├── __init__.py
│   ├── file_handler.py        # Validation, reading, metadata, sheet naming
│   ├── data_summary.py        # Business metric calculations
│   └── excel_export.py        # Compiling DataFrames into one workbook
│
├── requirements.txt
└── README.md
```

## Installation

1. Make sure you have Python 3.9 or later installed.
2. (Recommended) create a virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate      # Windows: venv\Scripts\activate
   ```
3. Install the dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Running the Application

From the `excel-compilation-tool` folder, run:

```bash
streamlit run app.py
```

Streamlit will open the app in your default browser (typically at
`http://localhost:8501`).

## How to Upload Files

1. Go to **Section 1 — Upload Files**.
2. Click the uploader and select one or more `.csv`, `.xlsx`, or `.xls`
   files (you can select multiple files at once, or drag-and-drop them).
3. Uploaded files are validated automatically. **Section 2 — Uploaded
   Files** shows a status table; any problem files are listed with a
   clear error message explaining what went wrong.

## How to Compile Files

1. Once at least one file has validated successfully, scroll to
   **Section 5 — Compile**.
2. Click **Compile Files**. The app builds one Excel workbook in memory,
   with each successfully validated file placed on its own worksheet.
3. A **Download Compiled Excel File** button appears — click it to save
   `Compiled_Data.xlsx` to your computer.

## Output Description

- The output is a single `.xlsx` workbook.
- Each uploaded file becomes one worksheet.
- Worksheet names are based on the original filename with the extension
  removed (e.g. `Portfolio Listing.xlsx` → `Portfolio Listing`).
- If a worksheet name would exceed Excel's 31-character limit, contain
  characters Excel disallows (`\ / ? * [ ] :`), or duplicate another sheet
  in the same workbook, it is adjusted as minimally as possible (truncated
  and/or given a `(2)`-style suffix) while staying as close as possible to
  the original filename. This only affects the *worksheet name* — never
  the data inside it.
- Column headers, row order, values, and formatting are preserved as
  uploaded; the app does not sort, filter, deduplicate, or clean the data.

## Troubleshooting

| Problem | Likely Cause | What To Do |
|---|---|---|
| "Unable to process `<file>`. The file appears to be empty or corrupted." | The file has 0 bytes, or pandas could not parse it as a valid CSV/Excel file. | Re-export the file from its source system and re-upload. |
| "Unsupported file format." | The file extension isn't `.csv`, `.xlsx`, or `.xls`. | Convert or re-save the file in a supported format. |
| A metric shows a warning instead of a number | The expected column (e.g. `Engagement Id`) isn't present in that file, using that exact name. | Check the column headers in the source file match the expected names exactly. |
| Two files end up with worksheet names like `Portfolio Listing` and `Portfolio Listing (2)` | Two uploaded files share the same base filename. | This is expected behavior — both files are compiled safely into separate sheets; rename the source files if you'd prefer distinct sheet names. |
| Compilation fails with an error | An unexpected issue occurred while writing the workbook (e.g. out of memory on very large files). | Try compiling fewer/smaller files at a time, or check the error message for specifics. |

## Explanation

**How the application works:** `app.py` renders the UI and coordinates
three utility modules. Uploaded files are handed to `file_handler.py`,
which validates and reads each one into a pandas DataFrame. Recognized
dataset types are passed to `data_summary.py` to compute dashboard
metrics. When the user clicks **Compile Files**, all validated DataFrames
are handed to `excel_export.py`, which writes them into one in-memory
Excel workbook (via `pandas.ExcelWriter` with the `openpyxl` engine) that
is then offered as a download.

**How files are validated:** Each file is checked for a supported
extension and non-zero size before pandas attempts to read it. Read
failures (corrupt files, malformed CSVs, unreadable Excel files) are
caught and turned into a plain-language error message identifying the
specific file — processing continues for the remaining files rather than
stopping the whole batch.

**How the summary metrics are calculated:** The app inspects each file's
name to detect whether it matches a known dataset type (Non-Conformance
Data or Portfolio Listing). If a required column is missing, the metric
is skipped with a warning rather than causing an error. Metrics use
`nunique()` for unique-project counts and `count()` for non-null
record/day counts — never `info()`, which is a diagnostic tool, not a
calculation method.

**How data integrity is preserved:** The DataFrame returned by
`read_uploaded_file()` is passed through the entire pipeline unmodified —
metric functions only *read* from it, and `compile_workbook()` writes it
straight to a worksheet with `index=False` (to avoid adding a pandas
index column that wasn't in the source) and no sorting, filtering, or
transformation of any kind.

**How the Excel workbook is generated:** `pandas.ExcelWriter` (backed by
`openpyxl`) writes each file's DataFrame to its own worksheet in a shared
`BytesIO` buffer, which is then served to the browser via Streamlit's
`st.download_button` — no temporary files are written to disk.
