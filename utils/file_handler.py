"""File handling utilities: validation, reading, and metadata extraction.

This module is responsible for safely reading uploaded CSV/Excel files
without modifying their contents, validating them, and extracting the
metadata used elsewhere in the application. No cleaning, transformation,
or normalization of the underlying data is ever performed here.

Header-row detection: business exports often have a title or description
line above the real table header (e.g. "Non-Conformance Report - Q3 2026"
in row 1, with the actual column headers in row 2). This module scans the
first few rows of each file to find the row that actually looks like a
header, rather than always assuming row 1. The choice can be overridden
manually per file (see reparse_with_header_row) if auto-detection guesses
wrong.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from io import BytesIO

import pandas as pd

SUPPORTED_EXTENSIONS = {".csv", ".xlsx", ".xls"}
EXCEL_SHEET_NAME_MAX_LENGTH = 31
EXCEL_INVALID_SHEET_CHARS = ["\\", "/", "?", "*", "[", "]", ":"]
HEADER_ROW_SCAN_LIMIT = 20  # how many leading rows to scan for the header


@dataclass
class ProcessedFile:
    """Holds a successfully read file and its metadata."""

    original_filename: str
    display_name: str  # filename without extension
    file_type: str  # "CSV" or "Excel"
    dataframe: pd.DataFrame
    sheet_name: str
    rows: int
    columns: int
    column_names: list[str]
    header_row: int  # 0-indexed row within the raw file used as the header
    source_bytes: bytes  # original, completely untouched file bytes


@dataclass
class FileError:
    """Holds information about a file that could not be processed."""

    original_filename: str
    reason: str


def get_file_extension(filename: str) -> str:
    """Return the lowercase file extension, including the leading dot."""
    return os.path.splitext(filename)[1].lower()


def strip_extension(filename: str) -> str:
    """Return the filename without its extension."""
    return os.path.splitext(filename)[0]


def is_supported_format(filename: str) -> bool:
    """Check whether the file extension is one this app supports."""
    return get_file_extension(filename) in SUPPORTED_EXTENSIONS


def sanitize_sheet_name(name: str, existing_names: set[str]) -> str:
    """Produce a valid, unique Excel worksheet name as close as possible
    to the original filename (extension already removed).

    Excel worksheet names cannot exceed 31 characters and cannot contain
    the characters: \\ / ? * [ ] :
    """
    sanitized = name
    for ch in EXCEL_INVALID_SHEET_CHARS:
        sanitized = sanitized.replace(ch, "-")

    sanitized = sanitized.strip()
    if not sanitized:
        sanitized = "Sheet"

    sanitized = sanitized[:EXCEL_SHEET_NAME_MAX_LENGTH]

    final_name = sanitized
    counter = 2
    existing_lower = {n.lower() for n in existing_names}
    while final_name.lower() in existing_lower:
        suffix = f" ({counter})"
        max_base_len = EXCEL_SHEET_NAME_MAX_LENGTH - len(suffix)
        final_name = f"{sanitized[:max_base_len]}{suffix}"
        counter += 1

    return final_name


def _read_raw_grid(raw_bytes: bytes, extension: str) -> pd.DataFrame:
    """Read the file as a raw grid with NO header interpretation, so every
    row - including any title/preamble rows above the real table - is kept
    as plain data for inspection.
    """
    buffer = BytesIO(raw_bytes)
    if extension == ".csv":
        return pd.read_csv(buffer, header=None, dtype=object)
    return pd.read_excel(buffer, sheet_name=0, header=None, dtype=object)


def _detect_header_row(raw: pd.DataFrame, scan_limit: int = HEADER_ROW_SCAN_LIMIT) -> int:
    """Guess which row holds the real column headers.

    Scans the first few rows and scores each one on how header-like it is:
    mostly filled across the row, mostly text (not numbers/dates), and
    mostly unique values (headers rarely repeat). A single description
    sentence in row 1 (which fills only one cell) scores far lower than
    the actual header row and is skipped automatically.
    """
    n_cols = raw.shape[1]
    limit = min(scan_limit, raw.shape[0])
    best_row = 0
    best_score = -1.0

    for i in range(limit):
        row = raw.iloc[i]
        non_null = int(row.notna().sum())
        if non_null == 0:
            continue

        fill_ratio = non_null / n_cols if n_cols else 0
        text_like = sum(isinstance(v, str) and v.strip() != "" for v in row)
        text_ratio = text_like / non_null if non_null else 0
        unique_ratio = row.dropna().nunique() / non_null if non_null else 0

        # A real header row is expected to span most of the columns and be
        # mostly text. A stray title/description line usually fills only
        # one or two cells, so it fails this check and is skipped.
        if fill_ratio < 0.5 or text_ratio < 0.5:
            continue

        score = (fill_ratio * 2) + text_ratio + unique_ratio
        if score > best_score:
            best_score = score
            best_row = i

    return best_row


def _finalize_dataframe(raw: pd.DataFrame, header_row: int) -> pd.DataFrame:
    """Slice the raw grid into a proper DataFrame using the chosen header row."""
    header_values = raw.iloc[header_row]
    data = raw.iloc[header_row + 1:].reset_index(drop=True)

    columns = []
    for i, v in enumerate(header_values):
        if pd.isna(v) or str(v).strip() == "":
            columns.append(f"Unnamed: {i}")
        else:
            columns.append(str(v).strip())
    data.columns = columns
    return data


def read_uploaded_file(
    raw_bytes: bytes, filename: str, header_row: int | None = None
) -> tuple[pd.DataFrame, int]:
    """Read a file's data, auto-detecting the header row unless one is given.

    Returns:
        (dataframe, header_row_used) - header_row_used is 0-indexed and
        refers to its position in the original file.

    Raises:
        ValueError: if the file cannot be parsed or is empty.
    """
    extension = get_file_extension(filename)

    try:
        raw = _read_raw_grid(raw_bytes, extension)
    except pd.errors.EmptyDataError as exc:
        raise ValueError("The file appears to be empty or corrupted.") from exc
    except pd.errors.ParserError as exc:
        raise ValueError(
            "The file could not be parsed. It may be corrupted or in an "
            "unexpected format."
        ) from exc
    except Exception as exc:  # noqa: BLE001
        raise ValueError(f"The file could not be read ({exc.__class__.__name__}).") from exc

    if raw is None or raw.empty:
        raise ValueError("The file appears to be empty or corrupted.")

    if header_row is None:
        resolved_header_row = _detect_header_row(raw)
    else:
        resolved_header_row = max(0, min(header_row, raw.shape[0] - 1))

    df = _finalize_dataframe(raw, resolved_header_row)
    if df.empty:
        raise ValueError("The file appears to be empty or corrupted.")

    return df, resolved_header_row


def process_uploaded_files(
    uploaded_files: list,
) -> tuple[list[ProcessedFile], list[FileError]]:
    """Validate and read a batch of uploaded files.

    Returns a tuple of (successfully processed files, errors). The source
    bytes of every file are kept untouched in memory (ProcessedFile.source_bytes)
    so that: (a) formatting can be preserved when compiling Excel files, and
    (b) a file can be re-read with a manually chosen header row if the
    auto-detected one is wrong.
    """
    processed: list[ProcessedFile] = []
    errors: list[FileError] = []
    seen_sheet_names: set[str] = set()

    for uploaded_file in uploaded_files:
        filename = uploaded_file.name

        if not is_supported_format(filename):
            errors.append(
                FileError(
                    original_filename=filename,
                    reason=(
                        "Unsupported file format. Only .csv, .xlsx and .xls "
                        "files are supported."
                    ),
                )
            )
            continue

        uploaded_file.seek(0)
        raw_bytes = uploaded_file.read()

        if not raw_bytes:
            errors.append(
                FileError(
                    original_filename=filename,
                    reason="The file appears to be empty or corrupted.",
                )
            )
            continue

        try:
            df, header_row = read_uploaded_file(raw_bytes, filename)
        except ValueError as exc:
            errors.append(FileError(original_filename=filename, reason=str(exc)))
            continue
        except Exception as exc:  # noqa: BLE001
            errors.append(
                FileError(
                    original_filename=filename,
                    reason=f"Unable to process this file ({exc.__class__.__name__}).",
                )
            )
            continue

        display_name = strip_extension(filename)
        sheet_name = sanitize_sheet_name(display_name, seen_sheet_names)
        seen_sheet_names.add(sheet_name)

        file_type = "CSV" if get_file_extension(filename) == ".csv" else "Excel"

        processed.append(
            ProcessedFile(
                original_filename=filename,
                display_name=display_name,
                file_type=file_type,
                dataframe=df,
                sheet_name=sheet_name,
                rows=df.shape[0],
                columns=df.shape[1],
                column_names=list(df.columns),
                header_row=header_row,
                source_bytes=raw_bytes,
            )
        )

    return processed, errors


def reparse_with_header_row(pf: ProcessedFile, header_row: int) -> ProcessedFile:
    """Re-read an already-processed file using a manually chosen header row.

    Used when auto-detection picks the wrong row (e.g. an unusual layout).
    The original file bytes are reused - nothing is re-uploaded or altered.
    """
    df, resolved_header_row = read_uploaded_file(
        pf.source_bytes, pf.original_filename, header_row=header_row
    )
    return ProcessedFile(
        original_filename=pf.original_filename,
        display_name=pf.display_name,
        file_type=pf.file_type,
        dataframe=df,
        sheet_name=pf.sheet_name,
        rows=df.shape[0],
        columns=df.shape[1],
        column_names=list(df.columns),
        header_row=resolved_header_row,
        source_bytes=pf.source_bytes,
    )
