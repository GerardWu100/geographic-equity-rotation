# Presentation

`Geographic_Equity_Rotation.pptx` is a nine-slide presentation designed for an
11- to 13-minute talk. It summarizes the corrected workflow and results in
`StrategyPipeline.ipynb`.

## Rebuild

The generator uses the shared `html2pptx` converter and local Node.js
presentation dependencies:

```bash
npm install
npx playwright install chromium
npm run build:deck
```

The script writes intermediate HTML slides to `presentation/build/` and the
final deck to `presentation/Geographic_Equity_Rotation.pptx`.

## Research interpretation

The deck uses the corrected pipeline's stored results. The sample is
exploratory because earlier notebooks informed the feature families and
20-session horizon. Out-of-fold predictions are therefore not equivalent to a
new, untouched final evaluation.
