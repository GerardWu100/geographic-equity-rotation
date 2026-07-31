# Geographic Equity Rotation

Research repository for a geographic equity rotation strategy.

Prepared for **Workshop in Mathematical Finance** in the **Master of Mathematical Finance** program.

## Notebook workflow

- `EDA.ipynb` contains the original exploratory data analysis.
- `Signals.ipynb` contains the original signal research and single-ETF rotation.
- `Allocation.ipynb` contains the original allocation experiments.
- `StrategyPipeline.ipynb` is the corrected combined workflow. It removes copied
  non-session rows, excludes the invalid Japanese volatility series, uses
  purged walk-forward validation and next-close execution, applies a fixed
  top-two allocation, and adds simple volatility targeting with cash.

## Presentation

`presentation/Geographic_Equity_Rotation.pptx` is a nine-slide summary designed
for an 11- to 13-minute presentation. Speaker notes provide pacing and the main
talk track. Rebuild instructions are in `presentation/README.md`.

Install the declared environment and open Jupyter from the project root:

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
