/**
 * generate.js – combined slide deck
 *
 * Structure
 * ──────────
 *   Cover slide
 *   Section divider  01 · Overview
 *   Bullet content   – What is PptxGenJS?
 *   Table content    – Feature Comparison
 *   Section divider  02 · Chart Showcase
 *   1.  Clustered Column Bar
 *   2.  Horizontal Bar
 *   3.  Stacked Column
 *   4.  100% Stacked Column
 *   5.  3D Bar
 *   6.  Line
 *   7.  Smooth Line
 *   8.  Area
 *   9.  Pie
 *  10.  Doughnut
 *  11.  Scatter
 *  12.  Bubble
 *  13.  Radar
 *  14.  Combo (Bar + Line, dual axis)
 *   Closing cover
 *
 * All slides use the token-driven colour/font scheme from tokens/tokens.js.
 */

"use strict";

const path = require("path");
const PptxGenJS = require("pptxgenjs");
const { C, FONT, PALETTE } = require(path.join(__dirname, "../tokens/tokens.js"));

const pptx = new PptxGenJS();
pptx.layout  = "LAYOUT_WIDE";
pptx.author  = "tokens-ppt";
pptx.title   = "Combined Deck";

// ─── Slide Masters ───────────────────────────────────────────────────────────

// Cover / closing
pptx.defineSlideMaster({
  title: "COVER_MASTER",
  background: { color: C.surface },
  objects: [
    { rect: { x: 0, y: 0,    w: "100%", h: 0.12, fill: { color: C.accentText  } } },
    { rect: { x: 0, y: 7.38, w: "100%", h: 0.12, fill: { color: C.accentText  } } },
    {
      text: {
        text: "CONFIDENTIAL",
        options: { x: 0, y: 7.1, w: "100%", h: 0.28, align: "center", color: C.onSurfaceLow, fontSize: 9, fontFace: FONT.ui },
      },
    },
    {
      placeholder: {
        options: { name: "cover_title", type: "title", x: 1.0, y: 2.2, w: 11.3, h: 1.4 },
        text: "Presentation Title",
      },
    },
    {
      placeholder: {
        options: { name: "cover_subtitle", type: "body", x: 1.0, y: 3.8, w: 11.3, h: 0.9 },
        text: "Subtitle / Date",
      },
    },
  ],
  slideNumber: { x: 12.8, y: 7.1, color: C.onSurfaceLow, fontSize: 9 },
});

// Section divider
pptx.defineSlideMaster({
  title: "SECTION_MASTER",
  background: { color: C.containerHigh },
  objects: [
    { rect: { x: 0, y: 0, w: 0.18, h: "100%", fill: { color: C.accentText } } },
    {
      placeholder: {
        options: { name: "section_label", type: "body",  x: 0.5, y: 0.6, w: 8.5, h: 0.5 },
        text: "SECTION",
      },
    },
    {
      placeholder: {
        options: { name: "section_title", type: "title", x: 0.5, y: 1.4, w: 9.0, h: 2.0 },
        text: "Section Title",
      },
    },
  ],
  slideNumber: { x: 12.8, y: 7.1, color: C.onSurfaceLow, fontSize: 9 },
});

// Content (text / table slides)
pptx.defineSlideMaster({
  title: "CONTENT_MASTER",
  background: { color: C.surface },
  objects: [
    { rect: { x: 0, y: 0,    w: "100%", h: 0.70, fill: { color: C.containerHigh } } },
    { rect: { x: 0, y: 0.70, w: "100%", h: 0.05, fill: { color: C.accentText   } } },
    { rect: { x: 0, y: 7.10, w: "100%", h: 0.40, fill: { color: C.containerHigh } } },
    {
      text: {
        text: "tokens-ppt  |  Confidential",
        options: { x: 0.3, y: 7.12, w: 6, h: 0.35, color: C.onSurfaceLow, fontSize: 9, fontFace: FONT.ui },
      },
    },
    {
      placeholder: {
        options: { name: "slide_title", type: "title", x: 0.3, y: 0.06, w: 12.4, h: 0.62 },
        text: "Slide Title",
      },
    },
    {
      placeholder: {
        options: { name: "slide_body", type: "body", x: 0.4, y: 1.0, w: 12.5, h: 5.9 },
        text: "Content goes here…",
      },
    },
  ],
  slideNumber: { x: 12.8, y: 7.12, color: C.onSurfaceLow, fontSize: 9 },
});

// Chart slides (no body placeholder – charts are placed manually)
pptx.defineSlideMaster({
  title: "CHART_MASTER",
  background: { color: C.surface },
  objects: [
    {
      text: {
        text: "tokens-ppt  ·  Chart Showcase",
        options: { x: 0.3, y: 7.12, w: 8, h: 0.35, color: C.onSurfaceLow, fontSize: 9, fontFace: FONT.ui },
      },
    },
    {
      placeholder: {
        options: { name: "slide_title", type: "title", x: 0.3, y: 0.06, w: 12.4, h: 0.57 },
        text: "Slide Title",
      },
    },
  ],
  slideNumber: { x: 12.6, y: 7.12, color: C.onSurfaceLow, fontSize: 9 },
});

// ─── Helpers ─────────────────────────────────────────────────────────────────

function addCoverSlide(title, subtitle) {
  const s = pptx.addSlide({ masterName: "COVER_MASTER" });
  s.addText(title,    { placeholder: "cover_title",    color: C.white,      fontSize: 36, bold: true,   fontFace: FONT.heading });
  s.addText(subtitle, { placeholder: "cover_subtitle", color: C.accentText, fontSize: 18, italic: true, fontFace: FONT.ui });
  return s;
}

function addSectionSlide(number, title) {
  const s = pptx.addSlide({ masterName: "SECTION_MASTER" });
  s.addText(number, { placeholder: "section_label", color: C.accentText, fontSize: 20, bold: true,   fontFace: FONT.ui });
  s.addText(title,  { placeholder: "section_title", color: C.white,      fontSize: 44, bold: true,   fontFace: FONT.heading });
  return s;
}

function addContentSlide(title) {
  const s = pptx.addSlide({ masterName: "CONTENT_MASTER" });
  s.addText(title, { placeholder: "slide_title", color: C.white, fontSize: 22, bold: true, fontFace: FONT.heading });
  return s;
}

function addChartSlide(title) {
  const s = pptx.addSlide({ masterName: "CHART_MASTER" });
  s.addText(title, { placeholder: "slide_title", color: C.white, fontSize: 20, bold: true, fontFace: FONT.heading });
  return s;
}

// Common chart area
const CA = { x: 0.4, y: 0.85, w: 12.5, h: 6.0 };

// Shared chart style defaults
const CS = {
  catAxisLabelColor:   C.onSurface, valAxisLabelColor:   C.onSurface,
  catAxisLabelFontFace: FONT.ui,    valAxisLabelFontFace: FONT.ui,
  legendFontSize: 12, legendColor: C.onSurface, legendFontFace: FONT.ui,
  titleFontSize: 14,  titleColor:  C.white,     titleFontFace: FONT.heading,
  dataLabelColor: C.white, dataLabelFontSize: 10, dataLabelFontFace: FONT.ui,
  valGridLine: { color: C.outline, style: "dash", size: 1 },
  chartArea: { fill: { color: C.surface } },
  plotArea: { fill: { color: C.surfaceVariant } },
};

// ─── Cover ───────────────────────────────────────────────────────────────────

addCoverSlide(
  "Combined Slide Deck",
  "Overview · Charts · May 2026"
);

// ─── Section 01 · Overview ───────────────────────────────────────────────────

addSectionSlide("01", "Overview");

// Bullet slide
{
  const s = addContentSlide("What is PptxGenJS?");
  s.addText(
    [
      { text: "JavaScript library to create PowerPoint files\n",          options: { bullet: true, fontSize: 18, fontFace: FONT.body } },
      { text: "Works in Node.js, browsers, and React Native\n",           options: { bullet: true, fontSize: 18, fontFace: FONT.body } },
      { text: "Supports charts, tables, images, shapes & more\n",         options: { bullet: true, fontSize: 18, fontFace: FONT.body } },
      { text: "Fully typed with TypeScript definitions\n",                 options: { bullet: true, fontSize: 18, fontFace: FONT.body } },
      { text: "MIT licensed — free for personal and commercial use",       options: { bullet: true, fontSize: 18, fontFace: FONT.body } },
    ],
    { placeholder: "slide_body", color: C.onSurface }
  );
}

// Table slide
{
  const s = addContentSlide("Feature Comparison");
  s.addTable(
    [
      [
        { text: "Feature",         options: { bold: true, color: C.white, fill: { color: C.containerHigh } } },
        { text: "PptxGenJS",       options: { bold: true, color: C.white, fill: { color: C.containerHigh } } },
        { text: "Other libs",      options: { bold: true, color: C.white, fill: { color: C.containerHigh } } },
      ],
      [{ text: "Charts",           options: { color: C.onSurface } }, { text: "✔", options: { color: C.success  } }, { text: "Partial", options: { color: C.onSurface } }],
      [{ text: "Tables",           options: { color: C.onSurface } }, { text: "✔", options: { color: C.success  } }, { text: "✔",       options: { color: C.success  } }],
      [{ text: "Slide Masters",    options: { color: C.onSurface } }, { text: "✔", options: { color: C.success  } }, { text: "✘",       options: { color: C.accentText } }],
      [{ text: "TypeScript",       options: { color: C.onSurface } }, { text: "✔", options: { color: C.success  } }, { text: "Partial", options: { color: C.onSurface } }],
      [{ text: "Browser Support",  options: { color: C.onSurface } }, { text: "✔", options: { color: C.success  } }, { text: "Node only", options: { color: C.onSurface } }],
    ],
    {
      x: 0.4, y: 1.1, w: 12.5,
      colW: [4.5, 4.0, 4.0],
      border: { type: "solid", color: C.outline },
      fontSize: 16,
      align: "center",
      valign: "middle",
      rowH: 0.55,
      fontFace: FONT.ui,
      fill: { color: C.surfaceVariant },
    }
  );
}

// ─── Section 02 · Chart Showcase ─────────────────────────────────────────────

addSectionSlide("02", "Chart Showcase");

// 1. Clustered Column
{
  const s = addChartSlide("1 · Clustered Column Chart");
  s.addChart(
    pptx.ChartType.bar,
    [
      { name: "Product A", labels: ["Q1","Q2","Q3","Q4"], values: [45, 62, 78, 90] },
      { name: "Product B", labels: ["Q1","Q2","Q3","Q4"], values: [30, 55, 60, 75] },
      { name: "Product C", labels: ["Q1","Q2","Q3","Q4"], values: [20, 38, 52, 68] },
    ],
    {
      ...CA, ...CS,
      barDir: "col", barGrouping: "clustered", barGapWidthPct: 50,
      chartColors: PALETTE,
      showLegend: true, legendPos: "b",
      showValue: true,
      showTitle: true, title: "Quarterly Revenue by Product ($K)",
    }
  );
}

// 2. Horizontal Bar
{
  const s = addChartSlide("2 · Horizontal Bar Chart");
  s.addChart(
    pptx.ChartType.bar,
    [
      { name: "Sales",  labels: ["Alice","Bob","Carol","Dave","Eve"], values: [92, 85, 78, 70, 65] },
      { name: "Target", labels: ["Alice","Bob","Carol","Dave","Eve"], values: [80, 80, 80, 80, 80] },
    ],
    {
      ...CA, ...CS,
      barDir: "bar", barGrouping: "clustered",
      chartColors: [C.primary, C.accent],
      showLegend: true, legendPos: "b",
      showValue: true, dataLabelPosition: "inEnd",
      showTitle: true, title: "Sales Rep Performance vs Target",
    }
  );
}

// 3. Stacked Column
{
  const s = addChartSlide("3 · Stacked Column Chart");
  s.addChart(
    pptx.ChartType.bar,
    [
      { name: "APAC", labels: ["Jan","Feb","Mar","Apr","May","Jun"], values: [20,25,28,30,35,40] },
      { name: "EMEA", labels: ["Jan","Feb","Mar","Apr","May","Jun"], values: [15,18,22,25,28,32] },
      { name: "AMER", labels: ["Jan","Feb","Mar","Apr","May","Jun"], values: [30,35,40,38,42,50] },
    ],
    {
      ...CA, ...CS,
      barDir: "col", barGrouping: "stacked",
      chartColors: [C.primary, C.base300, C.accent],
      showLegend: true, legendPos: "b",
      showValue: true,
      showTitle: true, title: "Monthly Revenue by Region ($M) – Stacked",
    }
  );
}

// 4. 100% Stacked Column
{
  const s = addChartSlide("4 · 100% Stacked Column Chart");
  s.addChart(
    pptx.ChartType.bar,
    [
      { name: "Direct",  labels: ["2021","2022","2023","2024","2025"], values: [40,35,30,28,25] },
      { name: "Partner", labels: ["2021","2022","2023","2024","2025"], values: [35,38,40,42,44] },
      { name: "Online",  labels: ["2021","2022","2023","2024","2025"], values: [25,27,30,30,31] },
    ],
    {
      ...CA, ...CS,
      barDir: "col", barGrouping: "percentStacked",
      chartColors: [C.accentText, C.success, C.primary],
      showLegend: true, legendPos: "b",
      showPercent: true,
      valAxisLabelFormatCode: "0%",
      showTitle: true, title: "Sales Channel Mix – 100% Stacked",
    }
  );
}

// 5. 3D Bar
{
  const s = addChartSlide("5 · 3D Bar Chart");
  s.addChart(
    pptx.ChartType.bar3d,
    [
      { name: "Widgets", labels: ["North","South","East","West"], values: [120, 95, 140, 110] },
      { name: "Gadgets", labels: ["North","South","East","West"], values: [ 85, 70, 100,  90] },
    ],
    {
      ...CA,
      barDir: "col", bar3DShape: "cylinder",
      chartColors: [C.primary, C.accentText],
      showLegend: true, legendPos: "b", legendFontSize: 12, legendColor: C.onSurface, legendFontFace: FONT.ui,
      showValue: true, dataLabelColor: C.white, dataLabelFontSize: 10, dataLabelFontFace: FONT.ui,
      catAxisLabelColor: C.onSurface, valAxisLabelColor: C.onSurface,
      catAxisLabelFontFace: FONT.ui,  valAxisLabelFontFace: FONT.ui,
      showTitle: true, title: "Regional Unit Sales – 3D Cylinder",
      titleColor: C.white, titleFontSize: 14, titleFontFace: FONT.heading,
      chartArea: { fill: { color: C.surface } },
    }
  );
}

// 6. Line
{
  const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const s = addChartSlide("6 · Line Chart");
  s.addChart(
    pptx.ChartType.line,
    [
      { name: "Revenue",  labels: months, values: [150,180,160,200,220,250,240,270,260,300,310,340] },
      { name: "Expenses", labels: months, values: [120,130,125,140,145,155,150,160,158,165,170,175] },
      { name: "Profit",   labels: months, values: [ 30, 50, 35, 60, 75, 95, 90,110,102,135,140,165] },
    ],
    {
      ...CA, ...CS,
      chartColors: [C.primary, C.accentText, C.success],
      lineSize: 2.5, lineDataSymbol: "circle", lineDataSymbolSize: 7,
      showLegend: true, legendPos: "b",
      catGridLine: { style: "none" },
      showTitle: true, title: "Full-Year P&L Trend ($K)",
    }
  );
}

// 7. Smooth Line
{
  const s = addChartSlide("7 · Smooth Line Chart");
  s.addChart(
    pptx.ChartType.line,
    [
      { name: "CPU %",    labels: ["00:00","04:00","08:00","12:00","16:00","20:00","24:00"], values: [12,10,45,70,80,55,20] },
      { name: "Memory %", labels: ["00:00","04:00","08:00","12:00","16:00","20:00","24:00"], values: [40,42,60,75,78,65,50] },
    ],
    {
      ...CA, ...CS,
      lineSmooth: true, lineSize: 2.5, lineDataSymbol: "none",
      chartColors: [C.accent, C.base300],
      showLegend: true, legendPos: "b",
      valAxisMaxVal: 100, valAxisLabelFormatCode: "0\"%\"",
      showTitle: true, title: "Server Resource Utilisation (24h)",
    }
  );
}

// 8. Area
{
  const wks = ["Wk1","Wk2","Wk3","Wk4","Wk5","Wk6","Wk7","Wk8"];
  const s = addChartSlide("8 · Area Chart");
  s.addChart(
    pptx.ChartType.area,
    [
      { name: "Mobile",  labels: wks, values: [500, 620, 580, 700, 750, 800, 870, 940] },
      { name: "Desktop", labels: wks, values: [900, 850, 820, 780, 740, 700, 660, 620] },
    ],
    {
      ...CA, ...CS,
      chartColors: [C.primary, C.accentText],
      chartColorsOpacity: 65,
      showLegend: true, legendPos: "b",
      showTitle: true, title: "Weekly Active Users by Platform",
    }
  );
}

// 9. Pie
{
  const s = addChartSlide("9 · Pie Chart");
  s.addChart(
    pptx.ChartType.pie,
    [
      {
        name: "Market Share",
        labels: ["Company A","Company B","Company C","Company D","Others"],
        values: [38, 26, 18, 10, 8],
      },
    ],
    {
      ...CA,
      chartColors: [C.primary, C.accentText, C.accent, C.base300, C.outline],
      showLegend: true, legendPos: "r", legendFontSize: 13, legendColor: C.onSurface, legendFontFace: FONT.ui,
      showPercent: true, dataLabelColor: C.white, dataLabelFontSize: 12, dataLabelFontBold: true, dataLabelFontFace: FONT.ui,
      dataLabelFormatCode: "0\"%\"",
      showTitle: true, title: "Global Market Share 2025",
      titleColor: C.white, titleFontSize: 14, titleFontFace: FONT.heading,
      chartArea: { fill: { color: C.surface } },
    }
  );
}

// 10. Doughnut
{
  const s = addChartSlide("10 · Doughnut Chart");
  s.addChart(
    pptx.ChartType.doughnut,
    [
      {
        name: "Budget Allocation",
        labels: ["R&D","Sales","Marketing","Operations","HR"],
        values: [30, 25, 20, 15, 10],
      },
    ],
    {
      ...CA,
      holeSize: 55,
      chartColors: [C.primary, C.accent, C.accentText, C.base300, C.success],
      showLegend: true, legendPos: "r", legendFontSize: 13, legendColor: C.onSurface, legendFontFace: FONT.ui,
      showPercent: true, dataLabelFontSize: 12, dataLabelColor: C.white, dataLabelFontBold: true, dataLabelFontFace: FONT.ui,
      showTitle: true, title: "Annual Budget Allocation",
      titleColor: C.white, titleFontSize: 14, titleFontFace: FONT.heading,
      chartArea: { fill: { color: C.surface } },
    }
  );
}

// 11. Scatter
{
  const s = addChartSlide("11 · Scatter Chart");
  s.addChart(
    pptx.ChartType.scatter,
    [
      { name: "Group A", values: [14, 18, 20, 25, 30, 15, 22, 28], sizes: [3, 4, 5, 6, 7, 3, 5, 6] },
      { name: "Group B", values: [10, 12, 16, 22, 28, 35, 18, 24], sizes: [4, 3, 5, 7, 8, 6, 4, 5] },
    ],
    {
      ...CA, ...CS,
      chartColors: [C.primary, C.accentText],
      lineSize: 0, lineDataSymbol: "circle", lineDataSymbolSize: 10,
      showLegend: true, legendPos: "b",
      catGridLine: { color: C.outline, style: "dash", size: 1 },
      showTitle: true, title: "Scatter: Engagement vs Conversion",
    }
  );
}

// 12. Bubble
{
  const s = addChartSlide("12 · Bubble Chart");
  s.addChart(
    pptx.ChartType.bubble,
    [
      { name: "Segment X", values: [10, 20, 30, 15, 25], sizes: [15, 25, 10, 30, 20] },
      { name: "Segment Y", values: [ 5, 18, 28, 12, 35], sizes: [20, 10, 35, 15, 25] },
    ],
    {
      ...CA, ...CS,
      chartColors: [C.base300, C.accent],
      chartColorsOpacity: 70,
      lineSize: 0,
      showLegend: true, legendPos: "b",
      catGridLine: { color: C.outline, style: "dash", size: 1 },
      showTitle: true, title: "Bubble: Market Size vs Growth",
    }
  );
}

// 13. Radar
{
  const skills = ["Communication","Leadership","Technical","Creativity","Teamwork","Delivery"];
  const s = addChartSlide("13 · Radar Chart");
  s.addChart(
    pptx.ChartType.radar,
    [
      { name: "Alice", labels: skills, values: [90, 75, 85, 70, 95, 80] },
      { name: "Bob",   labels: skills, values: [70, 85, 90, 65, 75, 88] },
    ],
    {
      ...CA,
      radarStyle: "marker",
      chartColors: [C.primary, C.accentText],
      chartColorsOpacity: 30,
      lineSize: 2, lineDataSymbol: "circle", lineDataSymbolSize: 7,
      showLegend: true, legendPos: "b", legendFontSize: 12, legendColor: C.onSurface, legendFontFace: FONT.ui,
      catAxisLabelColor: C.onSurface, valAxisLabelColor: C.onSurface,
      catAxisLabelFontFace: FONT.ui,  valAxisLabelFontFace: FONT.ui,
      showTitle: true, title: "Employee Skill Assessment Radar",
      titleColor: C.white, titleFontSize: 14, titleFontFace: FONT.heading,
      chartArea: { fill: { color: C.surface } },
    }
  );
}

// 14. Combo – Clustered Bar + Line with secondary Y-axis
{
  const qtrs = ["Q1 23","Q2 23","Q3 23","Q4 23","Q1 24","Q2 24","Q3 24","Q4 24"];
  const s = addChartSlide("14 · Combo Chart (Bar + Line, dual axis)");
  s.addChart(
    [
      {
        type: pptx.ChartType.bar,
        data: [
          { name: "Revenue ($M)",  labels: qtrs, values: [4.2, 5.1, 4.8, 6.3, 5.5, 6.8, 7.1, 8.2] },
          { name: "Expenses ($M)", labels: qtrs, values: [3.1, 3.5, 3.3, 4.0, 3.8, 4.2, 4.5, 4.8] },
        ],
        options: { barDir: "col", barGrouping: "clustered", chartColors: [C.primary, C.accentText], showValue: false },
      },
      {
        type: pptx.ChartType.line,
        data: [
          { name: "Growth %", labels: qtrs, values: [5, 8, 6, 12, 9, 14, 16, 20] },
        ],
        options: {
          secondaryValAxis: true, secondaryCatAxis: true,
          chartColors: [C.accent],
          lineSize: 2.5, lineDataSymbol: "diamond", lineDataSymbolSize: 8,
        },
      },
    ],
    {
      ...CA,
      valAxes: [
        { valAxisTitle: "Revenue / Expenses ($M)", showValAxisTitle: true, valAxisTitleColor: C.onSurface, valAxisLabelColor: C.onSurface, valAxisTitleFontFace: FONT.ui, valAxisLabelFontFace: FONT.ui },
        { valAxisTitle: "YoY Growth (%)",          showValAxisTitle: true, valAxisTitleColor: C.accent,   valAxisLabelColor: C.accent,   valAxisMajorUnit: 5, valAxisTitleFontFace: FONT.ui, valAxisLabelFontFace: FONT.ui },
      ],
      catAxes: [
        { catAxisTitle: "" },
        { catAxisHidden: true },
      ],
      showLegend: true, legendPos: "b", legendFontSize: 12, legendColor: C.onSurface, legendFontFace: FONT.ui,
      catAxisLabelColor: C.onSurface, catAxisLabelFontFace: FONT.ui,
      valGridLine: { color: C.outline, style: "dash", size: 1 },
      showTitle: true, title: "Revenue vs Expenses with Growth Rate",
      titleColor: C.white, titleFontSize: 14, titleFontFace: FONT.heading,
      chartArea: { fill: { color: C.surface } },
      plotArea: { fill: { color: C.surfaceVariant } },
    }
  );
}

// ─── Closing cover ───────────────────────────────────────────────────────────

addCoverSlide(
  "Thank You",
  "Questions?  reach@democorp.io"
);

// ─── Save ────────────────────────────────────────────────────────────────────

pptx
  .writeFile({ fileName: "slides/output.pptx" })
  .then(() => console.log("✅  slides/output.pptx created successfully! (19 slides)"))
  .catch((err) => console.error("Error:", err));
