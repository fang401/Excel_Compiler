"""Business metric calculations for recognized dataset types.

All calculations here are strictly read-only: they never modify the
underlying DataFrame. They only compute values for display in the
dashboard summary.
"""

from __future__ import annotations

from dataclasses import dataclass

import re

import pandas as pd

# Keywords used to detect which known dataset type a file represents,
# based on its filename. Matching is case-insensitive and substring-based
# so that minor naming variations (e.g. "Non-Comformance" vs
# "Non-Conformance") are still recognized.
NON_CONFORMANCE_KEYWORDS = ["non-comformance", "non-conformance", "non conformance"]
PORTFOLIO_KEYWORDS = ["portfolio listing", "portfolio"]

ENGAGEMENT_ID_COLUMN = "Engagement Id"
STATUS_REPORT_COLUMN = "Status Report Non-compliance"
FINANCIALS_COLUMN = "Financials Non-compliance"
WORKPLANS_COLUMN = "Workplans Non-compliance"


@dataclass
class MetricResult:
    """A single metric value, or a note explaining why it is unavailable."""

    label: str
    value: object
    missing_column: str | None = None


def detect_dataset_type(display_name: str) -> str | None:
    """Guess which known business dataset a file represents, from its name.

    Returns "non_conformance", "portfolio", or None if the file doesn't
    match a recognized dataset type.
    """
    name_lower = display_name.lower()
    if any(kw in name_lower for kw in NON_CONFORMANCE_KEYWORDS):
        return "non_conformance"
    if any(kw in name_lower for kw in PORTFOLIO_KEYWORDS):
        return "portfolio"
    return None


def _normalize_header(name: object) -> str:
    """Normalize a column header for comparison purposes only.

    Lower-cases, collapses internal whitespace, and strips leading/trailing
    whitespace. This is used purely to *find* the right column to read from
    - it never changes the DataFrame's actual column names or values.
    """
    text = str(name)
    text = re.sub(r"\s+", " ", text).strip().lower()
    return text


def find_matching_column(df: pd.DataFrame, expected_name: str) -> str | None:
    """Find the real column in df that corresponds to expected_name.

    Matching is tolerant of things that commonly differ between a hardcoded
    expected name and a real-world export: extra/irregular whitespace,
    different capitalization, and trailing qualifiers such as
    "(numbers in days)" appended to the header. Returns the *actual* column
    name as it appears in df, or None if nothing matches closely enough.
    """
    expected_norm = _normalize_header(expected_name)

    # 1. Exact match, ignoring case/whitespace differences.
    for col in df.columns:
        if _normalize_header(col) == expected_norm:
            return col

    # 2. The real header starts with the expected name, e.g.
    #    "Status Report Non-compliance (numbers in days)".
    for col in df.columns:
        col_norm = _normalize_header(col)
        if col_norm.startswith(expected_norm):
            return col

    # 3. The expected name appears anywhere within the real header.
    for col in df.columns:
        if expected_norm in _normalize_header(col):
            return col

    return None


def _has_meaningful_value(value: object) -> bool:
    """True if a cell holds an actual value, not a null or whitespace-only
    "blank" cell (e.g. a cell that visually looks empty in Excel but
    actually contains a stray space character).
    """
    if pd.isna(value):
        return False
    if isinstance(value, str) and value.strip() == "":
        return False
    return True


def _count_non_missing(df: pd.DataFrame, column: str) -> int:
    """Count records/days that have an actual value in the given column.

    Null cells and whitespace-only "blank" cells are both excluded.
    """
    return int(df[column].apply(_has_meaningful_value).sum())


def compute_non_conformance_metrics(df: pd.DataFrame) -> list[MetricResult]:
    """Compute the Non-Conformance Data dashboard metrics.

    - Non-Compliance Projects: number of unique Engagement Id values.
    - Status/Financials/Workplans Non-compliance: number of non-null
      records in each corresponding column.
    """
    results: list[MetricResult] = []

    engagement_col = find_matching_column(df, ENGAGEMENT_ID_COLUMN)
    if engagement_col is not None:
        results.append(
            MetricResult(
                label="Non-Compliance Projects",
                value=int(df[engagement_col].nunique()),
            )
        )
    else:
        results.append(
            MetricResult(
                label="Non-Compliance Projects",
                value=None,
                missing_column=ENGAGEMENT_ID_COLUMN,
            )
        )

    for label, expected_column in [
        ("Status Report Non-compliance", STATUS_REPORT_COLUMN),
        ("Financials Non-compliance", FINANCIALS_COLUMN),
        ("Workplans Non-compliance", WORKPLANS_COLUMN),
    ]:
        actual_column = find_matching_column(df, expected_column)
        if actual_column is not None:
            results.append(MetricResult(label=label, value=_count_non_missing(df, actual_column)))
        else:
            results.append(MetricResult(label=label, value=None, missing_column=expected_column))

    return results


def compute_portfolio_metrics(df: pd.DataFrame) -> list[MetricResult]:
    """Compute the Portfolio Listing dashboard metrics.

    - Total Projects: number of unique Engagement Id values.
    """
    engagement_col = find_matching_column(df, ENGAGEMENT_ID_COLUMN)
    if engagement_col is not None:
        return [
            MetricResult(
                label="Total Projects",
                value=int(df[engagement_col].nunique()),
            )
        ]
    return [
        MetricResult(
            label="Total Projects",
            value=None,
            missing_column=ENGAGEMENT_ID_COLUMN,
        )
    ]
