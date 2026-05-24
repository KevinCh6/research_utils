"""Reusable utilities for literature and paper workflows."""

__all__ = [
    "MergeStats",
    "merge_jcr_dataframe",
    "merge_jcr_into_workbook",
    "translate_dataframe",
    "translate_excel",
    "translate_text",
]


def __getattr__(name: str):
    if name in {"MergeStats", "merge_jcr_dataframe", "merge_jcr_into_workbook"}:
        from paper_utils.web_of_science import merge_jcr_xlsx

        return getattr(merge_jcr_xlsx, name)
    if name in {"translate_dataframe", "translate_excel", "translate_text"}:
        from paper_utils.web_of_science import translate_xls

        return getattr(translate_xls, name)
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
