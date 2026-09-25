"""Excel Data Compilation Tool.

A Streamlit application for uploading multiple CSV/Excel files, reviewing
their contents and summary metrics, and compiling them into a single
downloadable Excel workbook - one worksheet per uploaded file - without
altering any of the source data.
"""

from __future__ import annotations

import streamlit as st

from utils.data_summary import (
    MetricResult,
    compute_non_conformance_metrics,
    compute_portfolio_metrics,
    detect_dataset_type,
)
from utils.excel_export import compile_workbook
from utils.file_handler import FileError, ProcessedFile, process_uploaded_files

st.set_page_config(
    page_title="Excel Data Compilation Tool",
    page_icon="📊",
    layout="wide",
)

CUSTOM_CSS = """
<style>
    .main .block-container { padding-top: 2rem; max-width: 1200px; }
    .app-header {
        font-size: 1.9rem;
        font-weight: 700;
        color: #1a1a2e;
        margin-bottom: 0.1rem;
    }
    .app-subheader {
        color: #555b66;
        font-size: 1rem;
        margin-bottom: 1.6rem;
    }
    .section-title {
        font-size: 1.15rem;
        font-weight: 600;
        margin-top: 1.8rem;
        margin-bottom: 0.6rem;
        color: #1a1a2e;
        border-bottom: 1px solid #e6e8eb;
        padding-bottom: 0.4rem;
    }
    div[data-testid="stMetric"] {
        background-color: #f8f9fb;
        border: 1px solid #e6e8eb;
        border-radius: 10px;
        padding: 0.9rem 1.1rem;
    }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


def render_header() -> None:
    """Render the application title and short description."""
    st.markdown('<div class="app-header">📊 Excel Data Compilation Tool</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="app-subheader">Upload multiple CSV or Excel files, review the '
        "dataset summary, and compile them into a single Excel workbook with "
        "separate worksheets.</div>",
        unsafe_allow_html=True,
    )


def render_upload_section() -> list:
    """Render the multi-file uploader and return the uploaded files."""
    st.markdown('<div class="section-title">1. Upload Files</div>', unsafe_allow_html=True)
    uploaded_files = st.file_uploader(
        "Upload CSV or Excel files (.csv, .xlsx, .xls) — any combination, any number.",
        type=["csv", "xlsx", "xls"],
        accept_multiple_files=True,
        help="You do not need to upload every file type. The app works with whichever files you provide.",
    )
    return uploaded_files or []


def render_file_summary(
    processed_files: list[ProcessedFile], file_errors: list[FileError]
) -> None:
    """Render the uploaded-files summary table and any validation errors."""
    st.markdown('<div class="section-title">2. Uploaded Files</div>', unsafe_allow_html=True)

    if not processed_files and not file_errors:
        st.info("No files uploaded yet. Use the uploader above to get started.")
        return

    if processed_files:
        table_rows = [
            {
                "File Name": pf.original_filename,
                "File Type": pf.file_type,
                "Rows": f"{pf.rows:,}",
                "Columns": pf.columns,
                "Status": "✓ Ready",
            }
            for pf in processed_files
        ]
        st.dataframe(table_rows, use_container_width=True, hide_index=True)

    for err in file_errors:
        st.error(f"❌ Unable to process **{err.original_filename}**. {err.reason}")


def render_single_metric(metric: MetricResult, source_name: str) -> None:
    """Render one metric, or a warning if its source column is missing."""
    if metric.missing_column is not None:
        st.warning(
            f"⚠️ **{metric.missing_column}** was not found in **{source_name}**. "
            f"'{metric.label}' cannot be calculated for this file."
        )
        return
    display_value = f"{metric.value:,}" if isinstance(metric.value, int) else metric.value
    st.metric(label=metric.label, value=display_value)


def render_data_summary(processed_files: list[ProcessedFile]) -> None:
    """Render the business-metric dashboard for recognized dataset types."""
    st.markdown('<div class="section-title">3. Data Summary</div>', unsafe_allow_html=True)

    recognized = [
        (pf, detect_dataset_type(pf.display_name))
        for pf in processed_files
    ]
    recognized = [(pf, kind) for pf, kind in recognized if kind is not None]

    if not recognized:
        st.info(
            "Upload a file matching a recognized dataset (e.g. a Non-Conformance "
            "Data or Portfolio Listing file) to see business metrics here."
        )
        return

    for pf, kind in recognized:
        st.markdown(f"**{pf.display_name}**")
        if kind == "non_conformance":
            metrics = compute_non_conformance_metrics(pf.dataframe)
        else:
            metrics = compute_portfolio_metrics(pf.dataframe)

        cols = st.columns(len(metrics))
        for col, metric in zip(cols, metrics):
            with col:
                render_single_metric(metric, pf.original_filename)


def render_preview_section(processed_files: list[ProcessedFile]) -> None:
    """Render a file selector and a preview of the selected dataset."""
    st.markdown('<div class="section-title">4. Data Preview</div>', unsafe_allow_html=True)

    if not processed_files:
        st.info("Nothing to preview yet.")
        return

    options = {pf.original_filename: pf for pf in processed_files}
    selected_name = st.selectbox("Select a file to preview", list(options.keys()))
    pf = options[selected_name]

    st.caption(f"{pf.rows:,} rows × {pf.columns} columns")
    with st.expander("Column names", expanded=False):
        st.write(", ".join(str(c) for c in pf.column_names))

    preview_rows = min(10, pf.rows)
    st.dataframe(pf.dataframe.head(preview_rows), use_container_width=True)
    st.caption(f"Showing first {preview_rows} of {pf.rows:,} rows. Data is shown exactly as uploaded.")


def render_compile_section(processed_files: list[ProcessedFile]) -> None:
    """Render the compile button and, once compiled, the download button."""
    st.markdown('<div class="section-title">5. Compile</div>', unsafe_allow_html=True)

    if not processed_files:
        st.info("Upload at least one valid file before compiling.")
        return

    if st.button("🗂️ Compile Files", type="primary"):
        try:
            workbook_buffer = compile_workbook(processed_files)
        except RuntimeError as exc:
            st.error(f"❌ {exc}")
        else:
            st.session_state["compiled_workbook"] = workbook_buffer.getvalue()
            st.session_state["compiled_file_count"] = len(processed_files)
            st.success(f"✓ Successfully compiled {len(processed_files)} file(s) into one workbook.")

    if "compiled_workbook" in st.session_state:
        st.markdown('<div class="section-title">6. Download</div>', unsafe_allow_html=True)
        st.download_button(
            label="⬇️ Download Compiled Excel File",
            data=st.session_state["compiled_workbook"],
            file_name="Compiled_Data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


def main() -> None:
    """Application entry point."""
    render_header()
    uploaded_files = render_upload_section()

    processed_files: list[ProcessedFile] = []
    file_errors: list[FileError] = []

    if uploaded_files:
        processed_files, file_errors = process_uploaded_files(uploaded_files)

    render_file_summary(processed_files, file_errors)
    render_data_summary(processed_files)
    render_preview_section(processed_files)
    render_compile_section(processed_files)


if __name__ == "__main__":
    main()
