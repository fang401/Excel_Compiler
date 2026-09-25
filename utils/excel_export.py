"""Excel workbook compilation.

Combines multiple uploaded files into a single Excel workbook, one
worksheet per file. For files that originated as Excel (.xlsx/.xls), the
*original worksheet* is copied directly into the output workbook so the
compiled file looks exactly like the source. This includes:

- Cell values and per-cell styles (fonts, fills/colors, borders, number
  formats, alignment).
- Merged cells, column widths, row heights, and frozen panes.
- Excel "Table" objects (Insert > Table) and their table style, e.g. the
  blue/banded look Excel applies via a named table style rather than
  per-cell coloring.
- Conditional formatting rules.

For CSV files - which carry no formatting to begin with - the raw values
are written as plain cells.
"""

from __future__ import annotations

import copy as copy_module
from copy import copy
from io import BytesIO

import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook.workbook import Workbook


def _copy_cells_and_layout(source_ws: Worksheet, target_ws: Worksheet) -> None:
    """Copy every cell (value + style) plus merges/widths/heights/panes."""
    for row in source_ws.iter_rows():
        for cell in row:
            new_cell = target_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.border = copy(cell.border)
                new_cell.fill = copy(cell.fill)
                new_cell.number_format = cell.number_format
                new_cell.protection = copy(cell.protection)
                new_cell.alignment = copy(cell.alignment)

    for merged_range in source_ws.merged_cells.ranges:
        target_ws.merge_cells(str(merged_range))

    for col_letter, dim in source_ws.column_dimensions.items():
        if dim.width is not None:
            target_ws.column_dimensions[col_letter].width = dim.width

    for row_idx, dim in source_ws.row_dimensions.items():
        if dim.height is not None:
            target_ws.row_dimensions[row_idx].height = dim.height

    if source_ws.freeze_panes:
        target_ws.freeze_panes = source_ws.freeze_panes

    target_ws.sheet_view.showGridLines = source_ws.sheet_view.showGridLines


def _copy_tables(source_ws: Worksheet, target_ws: Worksheet, used_table_names: set[str]) -> None:
    """Copy Excel Table objects (Insert > Table) including their style.

    A table's blue/banded appearance normally comes from a named table
    style (e.g. "TableStyleMedium2") referenced by the table definition,
    not from per-cell coloring, so it must be copied explicitly. Table
    names must be unique across the whole workbook, so duplicates (Excel's
    common default "Table1" in several source files) are renamed.
    """
    for table in source_ws.tables.values():
        new_table = copy_module.deepcopy(table)

        base_name = new_table.name or "Table"
        candidate = base_name
        counter = 2
        while candidate in used_table_names:
            candidate = f"{base_name}_{counter}"
            counter += 1
        new_table.name = candidate
        new_table.displayName = candidate
        used_table_names.add(candidate)

        target_ws.add_table(new_table)


def _copy_conditional_formatting(source_ws: Worksheet, target_ws: Worksheet) -> None:
    """Copy conditional formatting rules (color scales, banded rules, etc.)."""
    for cf_range in source_ws.conditional_formatting:
        for rule in cf_range.rules:
            target_ws.conditional_formatting.add(str(cf_range.sqref), copy_module.deepcopy(rule))


def _copy_worksheet_with_style(
    source_ws: Worksheet, target_wb: Workbook, sheet_name: str, used_table_names: set[str]
) -> None:
    """Copy source_ws into a new worksheet named sheet_name in target_wb,
    preserving appearance as closely as technically possible.
    """
    target_ws = target_wb.create_sheet(title=sheet_name)
    _copy_cells_and_layout(source_ws, target_ws)
    _copy_tables(source_ws, target_ws, used_table_names)
    _copy_conditional_formatting(source_ws, target_ws)


def _write_dataframe_plain(target_wb: Workbook, sheet_name: str, df) -> None:
    """Write a DataFrame's values as plain cells (used for CSV-sourced files,
    which have no original formatting to preserve).
    """
    ws = target_wb.create_sheet(title=sheet_name)
    ws.append([str(c) for c in df.columns])
    for row in df.itertuples(index=False, name=None):
        ws.append(list(row))


def compile_workbook(processed_files: list) -> BytesIO:
    """Compile every processed file into one workbook, one sheet each.

    Args:
        processed_files: list of file_handler.ProcessedFile objects.

    Returns:
        A BytesIO buffer containing the compiled .xlsx workbook.

    Raises:
        RuntimeError: if the workbook cannot be created.
    """
    output_wb = openpyxl.Workbook()
    # Remove the default blank sheet Workbook() creates automatically.
    output_wb.remove(output_wb.active)
    used_table_names: set[str] = set()

    try:
        for pf in processed_files:
            if pf.file_type == "Excel":
                source_wb = openpyxl.load_workbook(
                    BytesIO(pf.source_bytes), data_only=False
                )
                source_ws = source_wb[source_wb.sheetnames[0]]
                _copy_worksheet_with_style(source_ws, output_wb, pf.sheet_name, used_table_names)
            else:
                _write_dataframe_plain(output_wb, pf.sheet_name, pf.dataframe)
    except Exception as exc:  # noqa: BLE001
        raise RuntimeError(f"Failed to create the compiled workbook: {exc}") from exc

    buffer = BytesIO()
    output_wb.save(buffer)
    buffer.seek(0)
    return buffer
