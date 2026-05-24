# research_utils
研究工具

## Web of Science Excel 工具

这些工具只面向 Web of Science 导出的 `savedrecs.xls/.xlsx` 工作簿。

翻译 WoS 导出表中的标题和摘要:

```bash
uv run wos-translate-xls path/to/savedrecs.xls -o path/to/savedrecs_translated.xlsx
```

合并 JCR 分区信息:

```bash
uv run wos-merge-jcr-xlsx data/web_of_science/jcr/XR2026-UTF8.csv path/to/savedrecs_translated.xlsx -o path/to/savedrecs_translated_jcr.xlsx
```

Python 调用:

```python
from paper_utils.web_of_science import merge_jcr_into_workbook, translate_excel

translate_excel("path/to/savedrecs.xls", "path/to/savedrecs_translated.xlsx")
merge_jcr_into_workbook(
    "data/web_of_science/jcr/XR2026-UTF8.csv",
    "path/to/savedrecs_translated.xlsx",
)
```
