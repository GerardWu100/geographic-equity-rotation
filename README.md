# Sector Rotation Strategy

Research repository for a sector / cross-asset rotation strategy.

Prepared for **Workshop in Mathematical Finance** in the **Master of Mathematical Finance** program.

## Credits

Alex Khadra, Matthew Antoniuk, and Gerard Wu.

## Data

- `data/raw/Index_ETF_Data.xlsx` is the unchanged source workbook.
- `data/processed/Index_ETF_Data.csv` is a UTF-8 CSV conversion of the workbook's
  only worksheet. It preserves row order, column positions, blank cells, text,
  and the exact numeric values stored by Excel.

Regenerate and verify the CSV from the project root:

```bash
uv run python scripts/convert_xlsx_to_csv.py
```
