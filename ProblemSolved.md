# Problem Solved — Data Ingestion Challenges

## Overview

There were not many problems while developing the solution. However, the key issue was reading the source data reliably given the inconsistent formats present across the transaction files. Three distinct variations had to be handled:

1. Some files contained an additional column (e.g. an `id` or sequence column) not present in all files.
2. Some CSV files used a semicolon (`;`) as the field separator instead of the standard comma (`,`), while others were standard comma-delimited or had already been opened and re-saved by Excel.
3. Although not observed in the actual dataset, the possibility of the header row not appearing on the first line was anticipated and handled defensively.

---

## Solution

The fix was implemented in two functions inside `src/engine.py`:

### 1. Separator Auto-Detection (`_read_file`)

For CSV files, `pandas.read_csv` was called with `sep=None` and `engine="python"`, which instructs pandas to "choose" the field separator automatically before parsing. This transparently handles both comma and semicolon-delimited files without requiring any manual inspection or pre-processing of the files.

```python
raw = pd.read_csv(filepath, header=None, sep=None, engine="python", dtype=str)
```

### 2. Dynamic Header Row Detection (`_find_header_row`)

All files — both CSV and Excel — are loaded initially with `header=None`, meaning pandas treats every row as plain data. A dedicated function, `_find_header_row`, then scans each row in sequence and checks whether its values contain all four required column names (`Client`, `Country`, `Currency`, `Transaction`). The index of the first matching row is returned and used as the true header, regardless of its position in the file.

```python
def _find_header_row(raw: pd.DataFrame) -> int | None:
    for i, row in raw.iterrows():
        if REQUIRED_COLS.issubset({str(v).strip() for v in row}):
            return i
    return None
```

The data is then sliced to start from the row immediately after the detected header, and only the four required columns are selected by name:

```python
raw.columns = [str(v).strip() for v in raw.iloc[header_row]]
df = raw.iloc[header_row + 1:].reset_index(drop=True)
df = df[list(REQUIRED_COLS)].copy()
```

### Result

Selecting columns **by name** after header detection automatically discards any extra columns (such as `id`) without needing to know their position or count in advance. Files where the header scan yields no match are skipped with a warning logged to the session log file, ensuring a single malformed file does not abort the entire run.

---
