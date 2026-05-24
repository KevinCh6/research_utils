# Web of Science JCR reference data

`XR2026-UTF8.csv` is the 2026 journal partition reference table used by
`paper_utils.web_of_science.merge_jcr_xlsx`.

The merge tool currently reads these columns:

- `Journal`: journal name used as the lookup key
- `大类新锐分区`: major-category partition copied into the workbook
- `Top`: top-journal marker copied into the workbook

The lookup matches `Journal` from this CSV to `Source Title` in a Web of
Science export workbook.

The CSV itself is a local reference dataset. If the repository was restored
from scratch and the file is missing, place it at:

```text
data/web_of_science/jcr/XR2026-UTF8.csv
```
