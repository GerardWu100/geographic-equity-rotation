/**
 * Build the Geographic Equity Rotation research presentation.
 *
 * Inputs
 * ------
 * StrategyPipeline.ipynb
 *     Source of the reported sample, signal, allocation, and backtest results.
 *
 * Outputs
 * -------
 * presentation/Geographic_Equity_Rotation.pptx
 *     Editable 16:9 PowerPoint deck with speaker notes.
 *
 * Assumptions
 * -----------
 * The numerical results match the last stored execution of StrategyPipeline.ipynb.
 * All percentages shown in the deck are rounded presentation values.
 */

const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const { chromium } = require("playwright");

const PROJECT_ROOT = path.resolve(__dirname, "..");
const BUILD_DIR = path.join(__dirname, "build");
const OUTPUT_FILE = path.join(__dirname, "Geographic_Equity_Rotation.pptx");

// The shared converter lives outside this project, so add the project's local
// dependencies to Node.js's global lookup paths before loading that module.
process.env.NODE_PATH = [
  path.join(PROJECT_ROOT, "node_modules"),
  process.env.NODE_PATH,
].filter(Boolean).join(path.delimiter);
require("module").Module._initPaths();

// The shared converter supplies only TMPDIR to Chromium. Preserve the existing
// process environment as well so locally installed fonts and libraries remain
// discoverable on Homebrew-based Linux systems.
const launchChromium = chromium.launch.bind(chromium);
chromium.launch = (options = {}) => launchChromium({
  ...options,
  env: { ...process.env, ...(options.env || {}) },
});

const html2pptx = require(
  "/home/ai4000/projects/skills-shared/pptx/scripts/html2pptx.js",
);

const COLORS = {
  ink: "253238",
  muted: "637176",
  paper: "F6F1E7",
  white: "FFFFFF",
  teal: "147D7B",
  tealLight: "D7E9E5",
  amber: "D8942C",
  amberLight: "F3E2C2",
  red: "B64A3A",
  redLight: "F1D8D1",
  line: "C9C2B7",
  gray: "DDE0DB",
};

const SLIDE_WIDTH_INCHES = 10;
const SLIDE_HEIGHT_INCHES = 5.625;

/**
 * Escape text for insertion into HTML.
 *
 * @param {string} value - Untrusted or punctuation-rich text.
 * @returns {string} HTML-safe text.
 */
function escapeHtml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

/**
 * Wrap slide content in the shared 16:9 theme.
 *
 * @param {string} content - HTML content placed inside the slide body.
 * @param {string} bodyClass - Optional body-level layout class.
 * @returns {string} Complete HTML document.
 */
function slideHtml(content, bodyClass = "") {
  return `<!doctype html>
<html>
<head>
<meta charset="utf-8">
<style>
  * { box-sizing: border-box; }
  html { background: #${COLORS.paper}; }
  body {
    width: 720pt; height: 405pt; margin: 0; padding: 0;
    display: flex; position: relative; overflow: hidden;
    background: #${COLORS.paper}; color: #${COLORS.ink};
    font-family: Arial, Helvetica, sans-serif;
  }
  .slide { width: 720pt; height: 405pt; padding: 24pt 30pt 22pt 30pt; position: relative; }
  .eyebrow { margin: 0 0 7pt 0; color: #${COLORS.teal}; font-size: 10pt; font-weight: bold; letter-spacing: 1.4pt; text-transform: uppercase; }
  h1 { margin: 0; color: #${COLORS.ink}; font-family: Georgia, 'Times New Roman', serif; font-size: 31pt; line-height: 1.05; font-weight: normal; }
  h2 { margin: 0; color: #${COLORS.ink}; font-family: Georgia, 'Times New Roman', serif; font-size: 24pt; line-height: 1.08; font-weight: normal; }
  h3 { margin: 0; color: #${COLORS.ink}; font-size: 14pt; line-height: 1.15; }
  p { margin: 0; font-size: 13pt; line-height: 1.25; }
  .subhead { margin-top: 8pt; color: #${COLORS.muted}; font-size: 13pt; line-height: 1.28; }
  .top-rule { position: absolute; left: 30pt; top: 13pt; width: 55pt; height: 4pt; background: #${COLORS.amber}; }
  .footer { position: absolute; left: 30pt; right: 30pt; bottom: 9pt; display: flex; justify-content: space-between; }
  .footer p { color: #${COLORS.muted}; font-size: 7.5pt; }
  .card { background: #${COLORS.white}; border: 1pt solid #${COLORS.line}; border-radius: 8pt; padding: 14pt; }
  .teal-card { background: #${COLORS.tealLight}; border: 1pt solid #${COLORS.teal}; border-radius: 8pt; padding: 14pt; }
  .amber-card { background: #${COLORS.amberLight}; border: 1pt solid #${COLORS.amber}; border-radius: 8pt; padding: 14pt; }
  .red-card { background: #${COLORS.redLight}; border: 1pt solid #${COLORS.red}; border-radius: 8pt; padding: 14pt; }
  .metric { color: #${COLORS.ink}; font-family: Georgia, 'Times New Roman', serif; font-size: 28pt; line-height: 1; }
  .metric-label { margin-top: 5pt; color: #${COLORS.muted}; font-size: 9.5pt; line-height: 1.15; }
  .small { color: #${COLORS.muted}; font-size: 10.5pt; line-height: 1.25; }
  .tag { background: #${COLORS.tealLight}; border-radius: 12pt; padding: 5pt 9pt; }
  .tag p { color: #${COLORS.teal}; font-size: 9pt; font-weight: bold; }
  .grid-3 { display: flex; gap: 12pt; }
  .grid-3 > div { flex: 1; }
  .grid-4 { display: flex; gap: 9pt; }
  .grid-4 > div { flex: 1; }
  .two-col { display: flex; gap: 18pt; }
  .two-col > div { flex: 1; }
  .placeholder { background: #${COLORS.white}; }
  ul { margin: 8pt 0 0 0; padding-left: 18pt; }
  li { margin: 0 0 7pt 0; color: #${COLORS.ink}; font-size: 12.5pt; line-height: 1.22; }
  .${bodyClass} {}
</style>
</head>
<body class="${bodyClass}">${content}</body>
</html>`;
}

/**
 * Write a source HTML slide used by html2pptx.
 *
 * @param {number} index - One-based slide number.
 * @param {string} html - Complete slide HTML.
 * @returns {string} Absolute path to the written HTML file.
 */
function writeSlide(index, html) {
  const filePath = path.join(BUILD_DIR, `slide-${String(index).padStart(2, "0")}.html`);
  fs.writeFileSync(filePath, html, "utf8");
  return filePath;
}

/**
 * Add consistent source and slide-number footer text.
 *
 * @param {number} slideNumber - One-based slide number.
 * @param {string} source - Brief source or interpretation label.
 * @returns {string} Footer HTML.
 */
function footer(slideNumber, source) {
  return `<div class="footer"><p>${escapeHtml(source)}</p><p>${slideNumber} / 9</p></div>`;
}

/**
 * Add speaker notes with a pacing cue.
 *
 * @param {object} slide - PptxGenJS slide.
 * @param {string} pacing - Expected speaking time.
 * @param {string[]} notes - Speaker-note paragraphs.
 * @returns {void}
 */
function addSpeakerNotes(slide, pacing, notes) {
  const text = [`Pacing: ${pacing}`, ...notes].join("\n\n");
  slide.addNotes(text);
}

/**
 * Build and save the complete presentation.
 *
 * @returns {Promise<void>} Resolves after the PowerPoint file is written.
 */
async function buildPresentation() {
  fs.mkdirSync(BUILD_DIR, { recursive: true });

  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_16x9";
  pptx.author = "Alex Khadra, Matthew Antoniuk, and Gerard Wu";
  pptx.subject = "Exploratory geographic equity rotation research";
  pptx.title = "Geographic Equity Rotation";
  pptx.company = "Master of Mathematical Finance";
  pptx.lang = "en-CA";
  pptx.theme = {
    headFontFace: "Georgia",
    bodyFontFace: "Arial",
    lang: "en-CA",
  };

  const slides = [];

  slides.push(slideHtml(`
    <div style="position:absolute; left:0; top:0; width:245pt; height:405pt; background:#${COLORS.teal};"></div>
    <div style="position:absolute; left:42pt; top:49pt; width:150pt; height:150pt; border:2pt solid #${COLORS.paper}; border-radius:50%;"></div>
    <div style="position:absolute; left:88pt; top:95pt; width:58pt; height:58pt; background:#${COLORS.amber}; border-radius:50%;"></div>
    <div style="position:absolute; left:50pt; top:244pt; width:150pt;"><p style="color:#${COLORS.white}; font-size:10pt; line-height:1.4;">SPY · United States<br>EZU · Eurozone<br>EWJ · Japan</p></div>
    <div style="position:absolute; left:286pt; top:76pt; width:386pt;">
      <p class="eyebrow">Workshop in Mathematical Finance</p>
      <h1 style="font-size:39pt;">Geographic<br>Equity Rotation</h1>
      <p class="subhead" style="margin-top:14pt; width:345pt;">Can currency-option and market signals improve regional ETF selection?</p>
      <div style="margin-top:32pt; width:85pt; height:4pt; background:#${COLORS.amber};"></div>
      <p style="margin-top:14pt; color:#${COLORS.muted}; font-size:10pt;">Alex Khadra · Matthew Antoniuk · Gerard Wu</p>
    </div>
    ${footer(1, "Exploratory research · corrected pipeline")}
  `));

  slides.push(slideHtml(`
    <div class="slide"><div class="top-rule"></div>
      <p class="eyebrow">01 · Research question</p>
      <h2>One global question, three liquid regional ETFs</h2>
      <p class="subhead">Forecast each region's next 20-session return, then rotate toward the strongest predictions.</p>
      <div class="grid-3" style="margin-top:22pt;">
        <div class="card"><p class="metric">SPY</p><h3 style="margin-top:9pt;">United States</h3><p class="small" style="margin-top:8pt;">Broad U.S. equity exposure<br>Local currency: U.S. dollar</p></div>
        <div class="card"><p class="metric">EZU</p><h3 style="margin-top:9pt;">Eurozone</h3><p class="small" style="margin-top:8pt;">Large and mid-cap euro-area equities<br>Local currency: euro</p></div>
        <div class="card"><p class="metric">EWJ</p><h3 style="margin-top:9pt;">Japan</h3><p class="small" style="margin-top:8pt;">Japanese equity exposure<br>Local currency: yen</p></div>
      </div>
      <div class="teal-card" style="margin-top:16pt; display:flex; align-items:center; gap:17pt;">
        <p style="font-family:Georgia,serif; font-size:20pt; color:#${COLORS.teal}; width:105pt;">Hypothesis</p>
        <p style="font-size:12pt;">Trailing market, volume, volatility, and currency-option signals contain enough cross-regional information to improve risk-adjusted allocation.</p>
      </div>
      ${footer(2, "Universe: SPY, EZU, EWJ")}
    </div>
  `));

  slides.push(slideHtml(`
    <div class="slide"><div class="top-rule"></div>
      <p class="eyebrow">02 · Research design</p>
      <h2>The corrected pipeline enforces a tradable timeline</h2>
      <div class="grid-4" style="margin-top:20pt;">
        <div class="card"><p class="metric">3,625</p><p class="metric-label">valid market sessions<br>2012-02-27 to 2026-07-28</p></div>
        <div class="card"><p class="metric">2,954</p><p class="metric-label">out-of-fold prediction dates</p></div>
        <div class="card"><p class="metric">148</p><p class="metric-label">20-session rebalances</p></div>
        <div class="amber-card"><p class="metric">137</p><p class="metric-label">copied non-session rows removed</p></div>
      </div>
      <div style="margin-top:28pt; display:flex; align-items:center; gap:8pt;">
        <div class="teal-card" style="width:132pt; text-align:center;"><p style="font-weight:bold;">Decision close</p><p class="small" style="margin-top:4pt;">features known at t</p></div>
        <div style="width:38pt; height:2pt; background:#${COLORS.amber};"></div>
        <div class="card" style="width:132pt; text-align:center;"><p style="font-weight:bold;">Next close</p><p class="small" style="margin-top:4pt;">eligible fill at t+1</p></div>
        <div style="width:38pt; height:2pt; background:#${COLORS.amber};"></div>
        <div class="card" style="width:145pt; text-align:center;"><p style="font-weight:bold;">Hold 20 sessions</p><p class="small" style="margin-top:4pt;">exit at t+21</p></div>
        <div style="width:38pt; height:2pt; background:#${COLORS.amber};"></div>
        <div class="card" style="width:132pt; text-align:center;"><p style="font-weight:bold;">Purge 21</p><p class="small" style="margin-top:4pt;">no label overlap</p></div>
      </div>
      <p class="small" style="margin-top:20pt;">The Japanese volatility index candidate is excluded after an unexplained 100-fold jump; EWJ therefore uses fewer features.</p>
      ${footer(3, "StrategyPipeline.ipynb · corrected timing and data cleaning")}
    </div>
  `));

  slides.push(slideHtml(`
    <div class="slide"><div class="top-rule"></div>
      <p class="eyebrow">03 · Signal engine</p>
      <h2>Signal strength is positive—but uneven across regions</h2>
      <div class="two-col" style="margin-top:17pt;">
        <div style="width:39%; flex:none;">
          <div class="card"><h3>Trailing feature families</h3>
            <ul><li>20- and 60-session momentum</li><li>price trend and relative volume</li><li>currency risk-reversal demand</li><li>VIX for SPY and V2X for EZU</li></ul>
          </div>
          <div class="teal-card" style="margin-top:12pt;"><p style="font-weight:bold;">Model</p><p class="small" style="margin-top:6pt;">One standardized Ridge regression per ETF, tuned inside purged expanding-window folds.</p></div>
        </div>
        <div style="flex:1;">
          <div id="ic-chart" class="placeholder" style="width:378pt; height:219pt;"></div>
          <p class="small" style="margin-top:7pt;">Rank information coefficient (IC) is the Spearman correlation between predicted and realized returns. Higher is better; zero means no monotonic ranking skill.</p>
        </div>
      </div>
      ${footer(4, "Mean rank IC across five outer folds")}
    </div>
  `));

  slides.push(slideHtml(`
    <div class="slide"><div class="top-rule"></div>
      <p class="eyebrow">04 · Portfolio construction</p>
      <h2>Use ranks for selection; use volatility only for exposure</h2>
      <div style="display:flex; align-items:stretch; gap:10pt; margin-top:25pt;">
        <div class="card" style="width:145pt;"><p class="metric">1</p><h3 style="margin-top:8pt;">Rank forecasts</h3><p class="small" style="margin-top:8pt;">Do not treat noisy return magnitudes as precise portfolio inputs.</p></div>
        <div style="align-self:center; width:24pt; height:2pt; background:#${COLORS.amber};"></div>
        <div class="card" style="width:145pt;"><p class="metric">2</p><h3 style="margin-top:8pt;">Select top two</h3><p class="small" style="margin-top:8pt;">Allocate 50% to each selected ETF before risk scaling.</p></div>
        <div style="align-self:center; width:24pt; height:2pt; background:#${COLORS.amber};"></div>
        <div class="teal-card" style="width:155pt;"><p class="metric">12%</p><h3 style="margin-top:8pt;">Volatility cap</h3><p class="small" style="margin-top:8pt;">Scale exposure down using a 63-session covariance estimate; never lever up.</p></div>
        <div style="align-self:center; width:24pt; height:2pt; background:#${COLORS.amber};"></div>
        <div class="amber-card" style="width:145pt;"><p class="metric">17.7%</p><h3 style="margin-top:8pt;">Average cash</h3><p class="small" style="margin-top:8pt;">Residual capital stays in cash, modeled at a 0% return.</p></div>
      </div>
      <div class="grid-3" style="margin-top:17pt;">
        <div><p class="small"><b>Rebalance:</b> every 20 valid sessions</p></div>
        <div><p class="small"><b>Costs:</b> 5 basis points per ETF side</p></div>
        <div><p class="small"><b>Accounting:</b> trade from drifted holdings</p></div>
      </div>
      ${footer(5, "Fixed rule; no allocation or risk-target grid search")}
    </div>
  `));

  slides.push(slideHtml(`
    <div class="slide"><div class="top-rule"></div>
      <p class="eyebrow">05 · Comparative backtest</p>
      <h2>The strategy works—but SPY remains the stronger benchmark</h2>
      <div class="two-col" style="margin-top:16pt;">
        <div style="flex:1;"><div id="cagr-chart" class="placeholder" style="width:405pt; height:230pt;"></div></div>
        <div style="width:215pt; flex:none; display:flex; flex-direction:column; gap:10pt;">
          <div class="teal-card"><p class="metric">7.75%</p><p class="metric-label">risk-managed compound annual growth rate (CAGR)</p></div>
          <div class="card"><p class="metric">11.96%</p><p class="metric-label">annualized volatility</p></div>
          <div class="red-card"><p style="font-size:15pt; font-weight:bold; color:#${COLORS.red};">−6.19 percentage points</p><p class="metric-label">CAGR gap versus SPY only</p></div>
        </div>
      </div>
      <p class="small" style="margin-top:8pt;">All strategies use the same next-close execution dates and transaction-cost model. SPY-only CAGR is 13.94%.</p>
      ${footer(6, "Net of modeled ETF costs · 2014-09-26 to 2026-06-08")}
    </div>
  `));

  slides.push(slideHtml(`
    <div class="slide"><div class="top-rule"></div>
      <p class="eyebrow">06 · Risk trade-off</p>
      <h2>The volatility layer controls risk, not forecasting skill</h2>
      <div class="two-col" style="margin-top:18pt;">
        <div style="flex:1;"><div id="risk-chart" class="placeholder" style="width:410pt; height:235pt;"></div></div>
        <div style="width:212pt; flex:none;">
          <div class="teal-card"><p class="metric">−3.0</p><p class="metric-label">percentage points of annualized volatility versus unscaled top-two</p></div>
          <div class="amber-card" style="margin-top:11pt;"><p class="metric">−0.28</p><p class="metric-label">percentage points of maximum-drawdown improvement</p></div>
          <p class="small" style="margin-top:12pt;">Risk fell materially, but drawdown barely changed and Sharpe ratio declined from 0.819 to 0.689.</p>
        </div>
      </div>
      ${footer(7, "Risk-managed top-two versus the same allocation without scaling")}
    </div>
  `));

  slides.push(slideHtml(`
    <div class="slide"><div class="top-rule"></div>
      <p class="eyebrow">07 · Robustness and limits</p>
      <h2>The finding is directionally stable, but not yet validated</h2>
      <div class="two-col" style="margin-top:20pt;">
        <div style="width:280pt; flex:none;">
          <div class="card"><h3>20 rebalance starting phases</h3>
            <div style="margin-top:17pt; display:flex; align-items:end; gap:22pt;">
              <div><p class="metric">7.07%</p><p class="metric-label">minimum CAGR</p></div>
              <div><p class="metric">8.08%</p><p class="metric-label">median CAGR</p></div>
              <div><p class="metric">8.66%</p><p class="metric-label">maximum CAGR</p></div>
            </div>
            <p class="small" style="margin-top:17pt;">Sharpe ratio spans 0.596 to 0.805. Results are not driven by one arbitrary starting offset.</p>
          </div>
          <div class="amber-card" style="margin-top:12pt;"><p style="font-weight:bold;">Oracle warning</p><p class="small" style="margin-top:5pt;">The perfect-foresight result uses future returns and is not tradable.</p></div>
        </div>
        <div class="red-card" style="flex:1;"><h3>What still limits the claim</h3>
          <ul><li>Feature families and horizon were explored on this same history.</li><li>PX_LAST adjustment and dividend treatment need vendor confirmation.</li><li>Cash is modeled at 0%; costs omit changing spreads and market impact.</li><li>Japanese volatility data were excluded, not repaired.</li><li>A genuinely untouched future period is still required.</li></ul>
        </div>
      </div>
      ${footer(8, "Exploratory evidence, not an investable validation")}
    </div>
  `));

  slides.push(slideHtml(`
    <div style="position:absolute; left:0; top:0; width:720pt; height:405pt; background:#${COLORS.ink};"></div>
    <div style="position:absolute; left:30pt; top:20pt; width:55pt; height:4pt; background:#${COLORS.amber};"></div>
    <div style="position:absolute; left:46pt; top:56pt; width:625pt;">
      <p class="eyebrow" style="color:#${COLORS.amber};">08 · Decision</p>
      <h1 style="color:#${COLORS.white}; font-size:34pt;">Promising signal research.<br>Not yet a benchmark-beating strategy.</h1>
      <div class="grid-3" style="margin-top:31pt;">
        <div style="border-top:2pt solid #${COLORS.teal}; padding-top:12pt;"><p style="color:#${COLORS.white}; font-size:20pt; font-family:Georgia,serif;">01</p><p style="margin-top:8pt; color:#${COLORS.white}; font-size:12pt;">SPY and EZU show modest out-of-fold ranking signal.</p></div>
        <div style="border-top:2pt solid #${COLORS.amber}; padding-top:12pt;"><p style="color:#${COLORS.white}; font-size:20pt; font-family:Georgia,serif;">02</p><p style="margin-top:8pt; color:#${COLORS.white}; font-size:12pt;">Volatility scaling lowers risk but does not improve the core selection edge.</p></div>
        <div style="border-top:2pt solid #${COLORS.red}; padding-top:12pt;"><p style="color:#${COLORS.white}; font-size:20pt; font-family:Georgia,serif;">03</p><p style="margin-top:8pt; color:#${COLORS.white}; font-size:12pt;">SPY-only wins on CAGR and Sharpe ratio in the corrected sample.</p></div>
      </div>
      <div style="margin-top:28pt; background:#${COLORS.teal}; border-radius:7pt; padding:14pt 17pt;"><p style="color:#${COLORS.white}; font-size:13pt;"><b>Next test:</b> freeze the specification, confirm total-return data, and evaluate only on new observations.</p></div>
    </div>
    <div class="footer"><p style="color:#${COLORS.gray};">Geographic Equity Rotation</p><p style="color:#${COLORS.gray};">9 / 9</p></div>
  `));

  for (let index = 0; index < slides.length; index += 1) {
    const htmlPath = writeSlide(index + 1, slides[index]);
    const result = await html2pptx(htmlPath, pptx, { tmpDir: BUILD_DIR });
    const slide = result.slide;
    slide.background = { color: COLORS.paper };

    if (index === 3) {
      const area = result.placeholders.find((item) => item.id === "ic-chart");
      slide.addChart(pptx.charts.BAR, [{
        name: "Mean rank IC",
        labels: ["SPY", "EZU", "EWJ"],
        values: [0.196, 0.182, 0.081],
      }], {
        ...area,
        barDir: "bar",
        catAxisLabelFontFace: "Arial",
        catAxisLabelFontSize: 12,
        valAxisLabelFontFace: "Arial",
        valAxisLabelFontSize: 9,
        showTitle: true,
        title: "Out-of-fold return-ranking skill",
        titleFontFace: "Georgia",
        titleFontSize: 15,
        showLegend: false,
        showValue: true,
        dataLabelPosition: "outEnd",
        dataLabelColor: COLORS.ink,
        dataLabelFormatCode: "0.000",
        chartColors: [COLORS.teal],
        showCatAxisTitle: true,
        catAxisTitle: "ETF",
        showValAxisTitle: true,
        valAxisTitle: "Mean rank IC",
        valAxisMinVal: 0,
        valAxisMaxVal: 0.25,
        valAxisMajorUnit: 0.05,
        showValue: true,
        showCatName: false,
        showValAxis: true,
        showCatAxis: true,
        showGridLines: false,
        border: { color: COLORS.line, pt: 1 },
      });
    }

    if (index === 5) {
      const area = result.placeholders.find((item) => item.id === "cagr-chart");
      slide.addChart(pptx.charts.BAR, [{
        name: "CAGR",
        labels: ["SPY only", "Single-asset rotation", "Top-two unscaled", "Equal weight", "Risk-managed top-two"],
        values: [13.94, 12.51, 11.65, 10.12, 7.75],
      }], {
        ...area,
        barDir: "bar",
        catAxisLabelFontFace: "Arial",
        catAxisLabelFontSize: 10,
        valAxisLabelFontFace: "Arial",
        valAxisLabelFontSize: 9,
        showTitle: true,
        title: "Compound annual growth rate",
        titleFontFace: "Georgia",
        titleFontSize: 15,
        showLegend: false,
        showValue: true,
        dataLabelPosition: "outEnd",
        dataLabelColor: COLORS.ink,
        dataLabelFormatCode: "0.00\"%\"",
        chartColors: [COLORS.teal],
        showCatAxisTitle: true,
        catAxisTitle: "Strategy",
        showValAxisTitle: true,
        valAxisTitle: "CAGR (%)",
        valAxisMinVal: 0,
        valAxisMaxVal: 16,
        valAxisMajorUnit: 4,
        showGridLines: false,
        border: { color: COLORS.line, pt: 1 },
      });
    }

    if (index === 6) {
      const area = result.placeholders.find((item) => item.id === "risk-chart");
      slide.addChart(pptx.charts.BAR, [
        {
          name: "Risk-managed top-two",
          labels: ["CAGR", "Annualized volatility", "Maximum drawdown"],
          values: [7.75, 11.96, 27.56],
        },
        {
          name: "Top-two unscaled",
          labels: ["CAGR", "Annualized volatility", "Maximum drawdown"],
          values: [11.65, 14.91, 27.85],
        },
      ], {
        ...area,
        barDir: "bar",
        grouping: "clustered",
        catAxisLabelFontFace: "Arial",
        catAxisLabelFontSize: 10,
        valAxisLabelFontFace: "Arial",
        valAxisLabelFontSize: 9,
        showTitle: true,
        title: "Return and risk comparison",
        titleFontFace: "Georgia",
        titleFontSize: 15,
        showLegend: true,
        legendPos: "b",
        legendFontFace: "Arial",
        legendFontSize: 9,
        chartColors: [COLORS.teal, COLORS.amber],
        showCatAxisTitle: true,
        catAxisTitle: "Metric",
        showValAxisTitle: true,
        valAxisTitle: "Percent (%)",
        valAxisMinVal: 0,
        valAxisMaxVal: 32,
        valAxisMajorUnit: 8,
        showGridLines: false,
        border: { color: COLORS.line, pt: 1 },
      });
    }

    const speakerNotes = [
      ["1:00", "Open with the decision problem: whether regional ETF selection can be improved with a compact, causal signal set.", "Preview the conclusion: there is some signal, but not enough evidence to beat the simplest U.S. benchmark."],
      ["1:10", "Define the three exchange-traded funds and the 20-session forecasting horizon.", "Stress that the task is relative selection among regions, not forecasting the global equity market direction."],
      ["1:30", "Walk left to right through the timeline. Features are observed after the decision close; execution waits until the next valid close.", "The 21-session purge prevents a training label from extending into the next test period."],
      ["1:30", "Define the rank information coefficient before discussing values.", "SPY and EZU are modestly positive across folds. EWJ is weaker and also lacks a trustworthy Japanese volatility input."],
      ["1:10", "Explain why ranks become equal weights: forecast ordering is more credible than forecast magnitude.", "The risk layer can reduce exposure and hold cash; it cannot create signal skill."],
      ["1:40", "Lead with the benchmark comparison. Every displayed strategy is net of the same modeled ETF trading costs.", "The risk-managed strategy compounded positively, but SPY only delivered the highest compound annual growth rate and Sharpe ratio."],
      ["1:20", "Compare the same top-two rule with and without scaling.", "Annualized volatility falls by about three percentage points, yet maximum drawdown barely improves and return falls materially."],
      ["1:40", "The rebalance-phase test reduces concern that one arbitrary starting date drives the answer.", "Still, this is not a fresh holdout because earlier research informed the feature families and horizon. Mention the price-adjustment and simplified-cost caveats."],
      ["1:00", "Close on three claims only: modest regional ranking signal, useful exposure control, and no benchmark-beating evidence.", "The next honest experiment is to freeze the pipeline and wait for genuinely new data."],
    ][index];
    addSpeakerNotes(slide, speakerNotes[0], speakerNotes.slice(1));
  }

  await pptx.writeFile({ fileName: OUTPUT_FILE });
  console.log(`Wrote ${OUTPUT_FILE}`);
  console.log(`Slide size: ${SLIDE_WIDTH_INCHES} x ${SLIDE_HEIGHT_INCHES} inches`);
}

buildPresentation().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
