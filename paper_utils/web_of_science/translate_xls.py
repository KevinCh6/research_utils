"""Translate title and abstract columns in Web of Science export workbooks.

The module exposes reusable functions and also provides a small command line
interface:

    python -m paper_utils.web_of_science.translate_xls input.xls -o output.xlsx
"""

from __future__ import annotations

import argparse
import sys
import time
from collections.abc import Callable, Mapping
from pathlib import Path
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    import pandas as pd

DEFAULT_COLUMN_MAP = {
    "Article Title": "Article Title (中文)",
    "Abstract": "Abstract (中文)",
}

ProgressCallback = Callable[[str], None]


def translate_text(
    text: Any,
    translator: Any,
    *,
    max_retries: int = 3,
    retry_base_seconds: float = 2.0,
    progress: ProgressCallback | None = None,
) -> str:
    """Translate one cell value with retries."""
    pd = _pandas()
    if pd.isna(text) or str(text).strip() == "":
        return ""

    text_str = str(text).strip()

    for attempt in range(max_retries):
        try:
            return translator.translate(text_str)
        except Exception as exc:  # noqa: BLE001 - translator may raise transport-specific errors.
            if attempt < max_retries - 1:
                wait_seconds = retry_base_seconds**attempt
                if progress:
                    progress(f"  翻译失败，{wait_seconds:g}秒后重试... ({exc})")
                time.sleep(wait_seconds)
            else:
                if progress:
                    progress(f"  翻译失败，已跳过: {exc}")
                return f"[翻译失败] {text_str[:50]}..."

    return ""


def translate_dataframe(
    df: pd.DataFrame,
    *,
    column_map: Mapping[str, str] | None = None,
    source: str = "auto",
    target: str = "zh-CN",
    delay_seconds: float = 0.5,
    max_retries: int = 3,
    translator: Any | None = None,
    progress: ProgressCallback | None = None,
) -> pd.DataFrame:
    """Return a copy of *df* with translated columns appended."""
    pd = _pandas()
    from deep_translator import GoogleTranslator

    columns = dict(column_map or DEFAULT_COLUMN_MAP)
    missing_columns = [column for column in columns if column not in df.columns]
    if missing_columns:
        joined = ", ".join(repr(column) for column in missing_columns)
        raise ValueError(f"Excel 中缺少列: {joined}")

    active_translator = translator or GoogleTranslator(source=source, target=target)
    result = df.copy()

    for source_column, output_column in columns.items():
        if progress:
            progress(f"开始翻译 '{source_column}' -> '{output_column}'")

        translated_values: list[str] = []
        total_rows = len(result)
        for index, text in enumerate(result[source_column], 1):
            preview = str(text)[:80] if pd.notna(text) else "(空)"
            if progress:
                progress(f"  [{index}/{total_rows}] {preview}...")
            translated_values.append(
                translate_text(
                    text,
                    active_translator,
                    max_retries=max_retries,
                    progress=progress,
                )
            )
            if delay_seconds > 0:
                time.sleep(delay_seconds)

        result[output_column] = translated_values

    return result


def translate_excel(
    input_file: str | Path,
    output_file: str | Path | None = None,
    *,
    column_map: Mapping[str, str] | None = None,
    source: str = "auto",
    target: str = "zh-CN",
    delay_seconds: float = 0.5,
    max_retries: int = 3,
    sheet_name: str | int = 0,
    progress: ProgressCallback | None = None,
) -> Path:
    """Translate selected columns in an Excel file and write the output file."""
    pd = _pandas()
    input_path = Path(input_file).expanduser()
    output_path = Path(output_file).expanduser() if output_file else _default_output_path(input_path)

    if progress:
        progress(f"正在读取文件: {input_path}")
    df = pd.read_excel(input_path, sheet_name=sheet_name)
    if progress:
        progress(f"数据行数: {len(df)}")
        progress(f"原始列名: {df.columns.tolist()}")

    translated = translate_dataframe(
        df,
        column_map=column_map,
        source=source,
        target=target,
        delay_seconds=delay_seconds,
        max_retries=max_retries,
        progress=progress,
    )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    if progress:
        progress(f"正在保存结果到: {output_path}")
    translated.to_excel(output_path, index=False, engine="openpyxl")
    return output_path


def parse_column_map(values: list[str] | None) -> dict[str, str]:
    """Parse CLI column specs in ``source=output`` form."""
    if not values:
        return dict(DEFAULT_COLUMN_MAP)

    parsed: dict[str, str] = {}
    for value in values:
        if "=" not in value:
            raise ValueError(f"列映射格式错误: {value!r}，应为 source=output")
        source, output = value.split("=", 1)
        source = source.strip()
        output = output.strip()
        if not source or not output:
            raise ValueError(f"列映射格式错误: {value!r}，source 和 output 不能为空")
        parsed[source] = output
    return parsed


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Translate selected columns in a Web of Science export workbook."
    )
    parser.add_argument("input_file", help="Input .xls/.xlsx file")
    parser.add_argument("-o", "--output-file", help="Output .xlsx file. Defaults to *_translated.xlsx")
    parser.add_argument(
        "-c",
        "--column",
        action="append",
        help="Column mapping in source=output form. Can be used multiple times.",
    )
    parser.add_argument("--source", default="auto", help="Source language for GoogleTranslator")
    parser.add_argument("--target", default="zh-CN", help="Target language for GoogleTranslator")
    parser.add_argument("--sheet-name", default=0, help="Sheet name or index to read. Defaults to first sheet.")
    parser.add_argument("--delay", type=float, default=0.5, help="Delay between requests in seconds")
    parser.add_argument("--max-retries", type=int, default=3, help="Retries per cell")
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_arg_parser()
    args = parser.parse_args(argv)

    sheet_name: str | int = args.sheet_name
    if isinstance(sheet_name, str) and sheet_name.isdigit():
        sheet_name = int(sheet_name)

    try:
        output_path = translate_excel(
            args.input_file,
            args.output_file,
            column_map=parse_column_map(args.column),
            source=args.source,
            target=args.target,
            delay_seconds=args.delay,
            max_retries=args.max_retries,
            sheet_name=sheet_name,
            progress=print,
        )
    except Exception as exc:  # noqa: BLE001 - CLI should report a clean message.
        print(f"错误: {exc}", file=sys.stderr)
        return 1

    print(f"完成: {output_path}")
    return 0


def _default_output_path(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}_translated.xlsx")


def _pandas() -> Any:
    try:
        import pandas as pd
    except ModuleNotFoundError as exc:
        raise ModuleNotFoundError("缺少依赖 pandas，请先安装项目依赖: uv sync 或 pip install -e .") from exc
    return pd


if __name__ == "__main__":
    raise SystemExit(main())
