const PptxGenJS = require('pptxgenjs');

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_4x3'; // 10" x 7.5"

// Brand colors
const RED     = 'CC2133';   // Halliburton Cardinal Red
const DGRAY   = '43494E';   // Mako Dark Gray
const LGRAY   = 'D9DCDE';   // Light gray for table alt rows
const WHITE   = 'FFFFFF';
const DIVIDER = 'CCCCCC';

const slide = pptx.addSlide();

// ─────────────────────────────────────────────
// BACKGROUND (white)
// ─────────────────────────────────────────────
slide.background = { color: WHITE };

// ─────────────────────────────────────────────
// TITLE BAR  (y=0.15)
// ─────────────────────────────────────────────
// Red accent strip on left
slide.addShape(pptx.ShapeType.rect, {
  x: 0.2, y: 0.10, w: 0.10, h: 0.40,
  fill: { color: RED }, line: { color: RED }
});

slide.addText('Halliburton Company (NYSE: HAL)', {
  x: 0.38, y: 0.10, w: 7.80, h: 0.40,
  fontSize: 20, bold: true, color: DGRAY, fontFace: 'Arial', valign: 'middle'
});

slide.addText('Oilfield Services  |  Energy', {
  x: 0.38, y: 0.50, w: 7.80, h: 0.22,
  fontSize: 9, color: '777777', fontFace: 'Arial', valign: 'top'
});

// Ticker badge
slide.addShape(pptx.ShapeType.rect, {
  x: 8.30, y: 0.10, w: 1.45, h: 0.40,
  fill: { color: RED }, line: { color: RED }
});
slide.addText('HAL  |  NYSE', {
  x: 8.30, y: 0.10, w: 1.45, h: 0.40,
  fontSize: 9.5, bold: true, color: WHITE, fontFace: 'Arial', align: 'center', valign: 'middle'
});

// Horizontal title separator line
slide.addShape(pptx.ShapeType.line, {
  x: 0.20, y: 0.76, w: 9.60, h: 0,
  line: { color: RED, width: 1.5 }
});

// ─────────────────────────────────────────────
// VERTICAL DIVIDER (center column separator)
// ─────────────────────────────────────────────
slide.addShape(pptx.ShapeType.line, {
  x: 5.10, y: 0.82, w: 0, h: 6.42,
  line: { color: DIVIDER, width: 0.75 }
});

// Horizontal mid divider (between top and bottom quadrants)
slide.addShape(pptx.ShapeType.line, {
  x: 0.20, y: 3.72, w: 9.60, h: 0,
  line: { color: DIVIDER, width: 0.75 }
});

// ─────────────────────────────────────────────
// QUADRANT 1: COMPANY OVERVIEW  (top-left)
// x=0.20, y=0.82, w=4.70
// ─────────────────────────────────────────────
// Accent bar
slide.addShape(pptx.ShapeType.rect, {
  x: 0.20, y: 0.84, w: 0.08, h: 0.24,
  fill: { color: RED }, line: { color: RED }
});
slide.addText('Company Overview', {
  x: 0.34, y: 0.82, w: 4.60, h: 0.28,
  fontSize: 11, bold: true, color: DGRAY, fontFace: 'Arial'
});

const overviewBullets = [
  { text: 'HQ: Houston, TX (dual HQ: Dubai, UAE); Founded: 1919', options: { bullet: { indent: 8 }, breakLine: true } },
  { text: 'Employees: ~46,000 globally in 70+ countries', options: { bullet: { indent: 8 }, breakLine: true } },
  { text: 'CEO: Jeff Miller (Chairman, President & CEO since 2019)', options: { bullet: { indent: 8 }, breakLine: true } },
  { text: 'CFO: Eric Carre (EVP & CFO since May 2022)', options: { bullet: { indent: 8 }, breakLine: true } },
  { text: 'Market Cap: ~$30.1B | Price: $35.93 | Shares: 837.5M', options: { bullet: { indent: 8 }, breakLine: true } },
  { text: 'P/E: ~23.9x | Dividend Yield: 1.9% | Beta: ~1.3', options: { bullet: { indent: 8 }, breakLine: true } },
  { text: '52-Wk Range: $18.72 – $37.03 (+40.9% YoY)', options: { bullet: { indent: 8 } } },
];
slide.addText(overviewBullets, {
  x: 0.22, y: 1.14, w: 4.65, h: 2.50,
  fontSize: 9, fontFace: 'Arial', color: '222222',
  valign: 'top', paraSpaceAfter: 4
});

// ─────────────────────────────────────────────
// QUADRANT 2: BUSINESS & POSITIONING  (top-right)
// x=5.20, y=0.82, w=4.60
// ─────────────────────────────────────────────
slide.addShape(pptx.ShapeType.rect, {
  x: 5.20, y: 0.84, w: 0.08, h: 0.24,
  fill: { color: RED }, line: { color: RED }
});
slide.addText('Business & Positioning', {
  x: 5.34, y: 0.82, w: 4.50, h: 0.28,
  fontSize: 11, bold: true, color: DGRAY, fontFace: 'Arial'
});

const bizBullets = [
  { text: "World's 2nd-largest oilfield services company; #1 in North America by market share", options: { bullet: { indent: 8 }, breakLine: true } },
  { text: 'Segments: Completion & Production (58% of rev); Drilling & Evaluation (42%)', options: { bullet: { indent: 8 }, breakLine: true } },
  { text: 'Leading position in hydraulic fracturing & completions (~50% of total revenue)', options: { bullet: { indent: 8 }, breakLine: true } },
  { text: 'Revenue by Region: North America $9.6B (42%); Middle East/Asia $6.1B (27%); LatAm $4.2B (18%); EMEA $3.0B (13%)', options: { bullet: { indent: 8 }, breakLine: true } },
  { text: 'Growth engines: drilling technology, unconventionals, well intervention, artificial lift', options: { bullet: { indent: 8 }, breakLine: true } },
  { text: 'Returned >$1.6B to shareholders in FY2024 (~60% free cash flow return rate)', options: { bullet: { indent: 8 }, breakLine: true } },
  { text: 'Customers: major IOCs, NOCs, independents across 70+ countries', options: { bullet: { indent: 8 } } },
];
slide.addText(bizBullets, {
  x: 5.22, y: 1.14, w: 4.55, h: 2.50,
  fontSize: 9, fontFace: 'Arial', color: '222222',
  valign: 'top', paraSpaceAfter: 4
});

// ─────────────────────────────────────────────
// QUADRANT 3: KEY FINANCIALS  (bottom-left)
// x=0.20, y=3.82, w=4.70
// ─────────────────────────────────────────────
slide.addShape(pptx.ShapeType.rect, {
  x: 0.20, y: 3.78, w: 0.08, h: 0.24,
  fill: { color: RED }, line: { color: RED }
});
slide.addText('Key Financials & Valuation', {
  x: 0.34, y: 3.76, w: 4.60, h: 0.28,
  fontSize: 11, bold: true, color: DGRAY, fontFace: 'Arial'
});

// Financial table header style
const hdr = { bold: true, fill: { color: RED }, color: WHITE, fontFace: 'Arial', fontSize: 9, valign: 'middle', align: 'center' };
const col0 = { fontFace: 'Arial', fontSize: 9, bold: false, color: '222222', valign: 'middle', align: 'left' };
const col1 = { fontFace: 'Arial', fontSize: 9, bold: false, color: '222222', valign: 'middle', align: 'center' };
const altFill = { fill: { color: 'F4F4F4' } };

const tableData = [
  [
    { text: 'Metric ($B unless noted)', options: { ...hdr, align: 'left' } },
    { text: 'FY2023A', options: hdr },
    { text: 'FY2024A', options: hdr },
    { text: 'FY2025E', options: hdr },
  ],
  [
    { text: 'Revenue', options: col0 },
    { text: '$23.0', options: col1 },
    { text: '$22.9', options: col1 },
    { text: '$22.2', options: col1 },
  ],
  [
    { text: 'YoY Growth', options: { ...col0, ...altFill } },
    { text: '+13.0%', options: { ...col1, ...altFill } },
    { text: '-0.4%', options: { ...col1, ...altFill } },
    { text: '-3.3%', options: { ...col1, ...altFill } },
  ],
  [
    { text: 'Operating Income', options: col0 },
    { text: '$4.1', options: col1 },
    { text: '$3.8', options: col1 },
    { text: '~$3.5', options: col1 },
  ],
  [
    { text: 'Op. Margin', options: { ...col0, ...altFill } },
    { text: '17.8%', options: { ...col1, ...altFill } },
    { text: '16.6%', options: { ...col1, ...altFill } },
    { text: '~15.8%', options: { ...col1, ...altFill } },
  ],
  [
    { text: 'EBITDA', options: col0 },
    { text: '~$4.5', options: col1 },
    { text: '~$4.3', options: col1 },
    { text: '$4.2', options: col1 },
  ],
  [
    { text: 'EBITDA Margin', options: { ...col0, ...altFill } },
    { text: '~19.6%', options: { ...col1, ...altFill } },
    { text: '~18.8%', options: { ...col1, ...altFill } },
    { text: '19.1%', options: { ...col1, ...altFill } },
  ],
  [
    { text: 'Net Income', options: col0 },
    { text: '$2.64', options: col1 },
    { text: '$2.50', options: col1 },
    { text: '$1.28', options: col1 },
  ],
  [
    { text: 'EPS (GAAP, diluted)', options: { ...col0, ...altFill } },
    { text: '$2.92', options: { ...col1, ...altFill } },
    { text: '$2.83', options: { ...col1, ...altFill } },
    { text: '~$1.50E', options: { ...col1, ...altFill } },
  ],
  [
    { text: 'Free Cash Flow', options: col0 },
    { text: '$2.3', options: col1 },
    { text: '$2.6', options: col1 },
    { text: '—', options: col1 },
  ],
  [
    { text: 'Enterprise Value / EBITDA', options: { ...col0, ...altFill } },
    { text: '—', options: { ...col1, ...altFill } },
    { text: '6.3x', options: { ...col1, ...altFill } },
    { text: '—', options: { ...col1, ...altFill } },
  ],
  [
    { text: 'Market Cap', options: col0 },
    { text: '—', options: col1 },
    { text: '—', options: col1 },
    { text: '$30.1B', options: col1 },
  ],
];

slide.addTable(tableData, {
  x: 0.22, y: 4.10, w: 4.65, h: 3.10,
  fontFace: 'Arial', fontSize: 9,
  border: { pt: 0.5, color: DIVIDER },
  rowH: 0.24,
  colW: [2.10, 0.90, 0.90, 0.90],
});

// ─────────────────────────────────────────────
// QUADRANT 4: STOCK PERFORMANCE & OWNERSHIP  (bottom-right)
// x=5.20, y=3.82, w=4.60
// ─────────────────────────────────────────────
slide.addShape(pptx.ShapeType.rect, {
  x: 5.20, y: 3.78, w: 0.08, h: 0.24,
  fill: { color: RED }, line: { color: RED }
});
slide.addText('Stock Performance & Ownership', {
  x: 5.34, y: 3.76, w: 4.50, h: 0.28,
  fontSize: 11, bold: true, color: DGRAY, fontFace: 'Arial'
});

// Stock chart (1Y bar approximation using line chart with monthly close data)
// Approximate monthly closing prices Mar 2025 – Mar 2026 (HAL)
slide.addChart(pptx.ChartType.line, [
  {
    name: 'HAL Price (USD)',
    labels: ["Mar'25","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec","Jan'26","Feb","Mar'26"],
    values: [24.71, 22.80, 21.50, 20.90, 22.30, 25.10, 27.80, 30.20, 33.50, 34.80, 36.20, 35.50, 35.93]
  }
], {
  x: 5.22, y: 4.06, w: 4.55, h: 2.00,
  chartColors: [RED],
  lineSmooth: true,
  showLegend: false,
  showValue: false,
  catAxisLabelFontSize: 8,
  valAxisLabelFontSize: 8,
  catAxisLabelColor: '555555',
  valAxisLabelColor: '555555',
  valAxisMinVal: 18,
  valAxisMaxVal: 40,
  valAxisMajorUnit: 5,
  dataLabelFontSize: 8,
  title: '1-Year Stock Performance (USD)',
  titleFontSize: 9,
  titleBold: false,
  titleColor: '555555',
  plotArea: { fill: { color: WHITE } },
  chartArea: { fill: { color: WHITE }, border: { color: WHITE } },
  catGridLine: { style: 'none' },
  valGridLine: { color: DIVIDER, style: 'solid', size: 0.5 },
});

// Top Shareholders sub-header
slide.addShape(pptx.ShapeType.rect, {
  x: 5.20, y: 6.14, w: 0.06, h: 0.20,
  fill: { color: RED }, line: { color: RED }
});
slide.addText('Top Shareholders', {
  x: 5.30, y: 6.13, w: 4.45, h: 0.22,
  fontSize: 9.5, bold: true, color: DGRAY, fontFace: 'Arial'
});

const shareholders = [
  { text: 'Vanguard Group:  ~8.8%', options: { bullet: { indent: 8 }, breakLine: true } },
  { text: 'BlackRock, Inc.:  ~7.5%', options: { bullet: { indent: 8 }, breakLine: true } },
  { text: 'Capital Research Global:  ~6.2%', options: { bullet: { indent: 8 }, breakLine: true } },
  { text: 'State Street Corp.:  ~4.4%', options: { bullet: { indent: 8 }, breakLine: true } },
  { text: 'Geode Capital Mgmt.:  ~2.1%', options: { bullet: { indent: 8 } } },
];
slide.addText(shareholders, {
  x: 5.22, y: 6.37, w: 4.55, h: 0.82,
  fontSize: 8.5, fontFace: 'Arial', color: '222222',
  valign: 'top', paraSpaceAfter: 2
});

// ─────────────────────────────────────────────
// FOOTER  (y=7.28)
// ─────────────────────────────────────────────
slide.addShape(pptx.ShapeType.line, {
  x: 0.20, y: 7.26, w: 9.60, h: 0,
  line: { color: RED, width: 0.75 }
});

slide.addText(
  'Sources: Halliburton Q4 2024 Earnings Release; SEC EDGAR 10-K FY2024; Bullfincher; StockAnalysis; MacroTrends; Fintel; BrandPalettes. As of March 13, 2026.',
  {
    x: 0.22, y: 7.28, w: 9.56, h: 0.20,
    fontSize: 6.5, fontFace: 'Arial', color: '888888', valign: 'middle'
  }
);

// ─────────────────────────────────────────────
// SAVE
// ─────────────────────────────────────────────
pptx.writeFile({ fileName: 'HAL_Strip_Profile.pptx' })
  .then(() => console.log('Saved: HAL_Strip_Profile.pptx'))
  .catch(e => { console.error(e); process.exit(1); });
