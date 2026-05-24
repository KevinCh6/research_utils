"""Merge JCR metadata into a Web of Science export workbook.

The module exposes ``merge_jcr_into_workbook`` for reuse and also supports:

    python -m paper_utils.web_of_science.merge_jcr_xlsx jcr.csv savedrecs_translated.xlsx
"""

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    import pandas as pd

DEFAULT_SHEET_NAME = "筛选后"
DEFAULT_JCR_COLUMNS = ("大类新锐分区", "Top")


@dataclass(frozen=True)
class MergeStats:
    """Summary of a JCR merge operation."""

    output_file: Path
    sheet_name: str
    row_count: int
    matched_count: int


def merge_jcr_dataframe(
    target: pd.DataFrame,
    jcr: pd.DataFrame,
    *,
    source_title_column: str = "Source Title",
    jcr_journal_column: str = "Journal",
    jcr_columns: tuple[str, ...] = DEFAULT_JCR_COLUMNS,
) -> tuple[pd.DataFrame, int]:
    """Return a copy of *target* with JCR columns merged by journal title."""
    required_target_columns = [source_title_column]
    required_jcr_columns = [jcr_journal_column, *jcr_columns]
    _require_columns(target, required_target_columns, "目标 Excel sheet")
    _require_columns(jcr, required_jcr_columns, "JCR CSV")

    normalized_jcr = jcr.loc[:, required_jcr_columns].copy()
    normalized_jcr[jcr_journal_column] = normalized_jcr[jcr_journal_column].astype(str).str.strip()
    normalized_jcr = normalized_jcr.drop_duplicates(subset=[jcr_journal_column], keep="first")

    merged = target.copy()
    key = merged[source_title_column].astype(str).str.strip()
    lookup = normalized_jcr.set_index(jcr_journal_column)
    for column in jcr_columns:
        merged[column] = key.map(lookup[column])

    matched_count = merged[jcr_columns[0]].notna().sum() if jcr_columns else 0
    return merged, int(matched_count)


def merge_jcr_into_workbook(
    jcr_csv: str | Path,
    workbook_in: str | Path,
    workbook_out: str | Path | None = None,
    *,
    sheet_name: str = DEFAULT_SHEET_NAME,
    source_title_column: str = "Source Title",
    jcr_journal_column: str = "Journal",
    jcr_columns: tuple[str, ...] = DEFAULT_JCR_COLUMNS,
) -> MergeStats:
    """Merge JCR columns into one sheet and write all workbook sheets back out."""
    pd = _pandas()
    jcr_path = Path(jcr_csv).expanduser()
    input_path = Path(workbook_in).expanduser()
    output_path = Path(workbook_out).expanduser() if workbook_out else _default_output_path(input_path)

    jcr = pd.read_csv(jcr_path, usecols=[jcr_journal_column, *jcr_columns])
    sheets = pd.read_excel(input_path, sheet_name=None)
    if sheet_name not in sheets:
        available = ", ".join(repr(name) for name in sheets)
        raise ValueError(f"未找到 sheet {sheet_name!r}，可用 sheet: {available}")

    sheets[sheet_name], matched_count = merge_jcr_dataframe(
        sheets[sheet_name],
        jcr,
        source_title_column=source_title_column,
        jcr_journal_column=jcr_journal_column,
        jcr_columns=jcr_columns,
    )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for name, sheet_df in sheets.items():
            sheet_df.to_excel(writer, sheet_name=name, index=False)

    return MergeStats(
        output_file=output_path,
        sheet_name=sheet_name,
        row_count=len(sheets[sheet_name]),
        matched_count=matched_count,
    )


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Merge JCR metadata into a Web of Science export workbook sheet."
    )
    parser.add_argument("jcr_csv", help="JCR CSV path")
    parser.add_argument("workbook_in", help="Input workbook path")
    parser.add_argument("-o", "--workbook-out", help="Output workbook path. Defaults to *_jcr.xlsx")
    parser.add_argument("--sheet-name", default=DEFAULT_SHEET_NAME, help="Sheet to update")
    parser.add_argument("--source-title-column", default="Source Title", help="Workbook journal title column")
    parser.add_argument("--jcr-journal-column", default="Journal", help="JCR CSV journal title column")
    parser.add_argument(
        "--jcr-column",
        action="append",
        help="JCR metadata column to copy. Can be used multiple times.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_arg_parser()
    args = parser.parse_args(argv)
    jcr_columns = tuple(args.jcr_column) if args.jcr_column else DEFAULT_JCR_COLUMNS

    try:
        stats = merge_jcr_into_workbook(
            args.jcr_csv,
            args.workbook_in,
            args.workbook_out,
            sheet_name=args.sheet_name,
            source_title_column=args.source_title_column,
            jcr_journal_column=args.jcr_journal_column,
            jcr_columns=jcr_columns,
        )
    except Exception as exc:  # noqa: BLE001 - CLI should report a clean message.
        print(f"错误: {exc}", file=sys.stderr)
        return 1

    print(f"写入: {stats.output_file}")
    print(f"「{stats.sheet_name}」行数: {stats.row_count}，匹配到 JCR 行数: {stats.matched_count}")
    return 0


def _default_output_path(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}_jcr.xlsx")


def _require_columns(df: pd.DataFrame, columns: list[str], label: str) -> None:
    missing_columns = [column for column in columns if column not in df.columns]
    if missing_columns:
        joined = ", ".join(repr(column) for column in missing_columns)
        raise ValueError(f"{label} 缺少列: {joined}")


def _pandas() -> Any:
    try:
        import pandas as pd
    except ModuleNotFoundError as exc:
        raise ModuleNotFoundError("缺少依赖 pandas，请先安装项目依赖: uv sync 或 pip install -e .") from exc
    return pd


if __name__ == "__main__":
    raise SystemExit(main())
