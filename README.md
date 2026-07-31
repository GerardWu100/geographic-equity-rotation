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
talk track. A PDF export is also in `presentation/`.

Install the declared environment and open Jupyter from the project root:

```bash
uv sync
uv run jupyter lab
```

## Credits

Alex Khadra, Matthew Antoniuk, and Gerard Wu.

## Data

- `data/raw/Index_ETF_Data.xlsx` and `data/raw/RiskReversalData.xlsx` are the
  unchanged source workbooks.
- `data/processed/Index_ETF_Data.csv` and `data/processed/RiskReversalData.csv`
  are UTF-8 CSV conversions of each workbook's only worksheet. They preserve
  row order, column positions, blank cells, text, and the exact numeric values
  stored by Excel.
- `data/Index_ETF_Data_Clean.csv` and `data/RiskReversalData_Clean.csv` are the
  cleaned series used by the notebooks.
- `data/processed/signals.csv` stores derived signal outputs.
