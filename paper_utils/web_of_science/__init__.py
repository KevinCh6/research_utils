"""Utilities for Web of Science export workbooks."""

from paper_utils.web_of_science.merge_jcr_xlsx import (
    MergeStats,
    merge_jcr_dataframe,
    merge_jcr_into_workbook,
)
from paper_utils.web_of_science.translate_xls import translate_dataframe, translate_excel, translate_text

__all__ = [
    "MergeStats",
    "merge_jcr_dataframe",
    "merge_jcr_into_workbook",
    "translate_dataframe",
    "translate_excel",
    "translate_text",
]
