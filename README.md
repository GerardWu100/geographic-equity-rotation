# Sector Rotation Strategy

Research repository for a sector / cross-asset rotation strategy.

Prepared for **Workshop in Mathematical Finance** in the **Master of Mathematical Finance** program.

## Notebook workflow

Run the notebooks from the project root in this order:

1. `EDA.ipynb` explores the cleaned market data.
2. `Signals.ipynb` engineers factors, fits walk-forward return models, and
   evaluates single-ETF rotation.
3. `Allocation.ipynb` studies several portfolio-allocation rules and documents
   why tuning those rules is unstable in this sample.
4. `PortfolioConstruction.ipynb` reuses the allocation notebook and implements
   the simplest portfolio extension: hold the two highest-predicted ETFs at
   50% each instead of buying only one.

Install the declared environment and open Jupyter:

```bash
uv sync
uv run jupyter lab
```

## Credits

Alex Khadra, Matthew Antoniuk, and Gerard Wu.

## Data

- `data/raw/Index_ETF_Data.xlsx` is the unchanged source workbook.
- `data/processed/Index_ETF_Data.csv` is a UTF-8 CSV conversion of the workbook's
  only worksheet. It preserves row order, column positions, blank cells, text,
  and the exact numeric values stored by Excel.
- `data/raw/RiskReversalData.xlsx` is the unchanged source workbook.
- `data/processed/RiskReversalData.csv` is a UTF-8 CSV conversion of the workbook's
  only worksheet, produced the same way: row order, column positions, blank
  cells, text, and the exact numeric values stored by Excel are preserved.

Regenerate and verify the CSV from the project root:

```bash
uv run python scripts/convert_xlsx_to_csv.py
```
