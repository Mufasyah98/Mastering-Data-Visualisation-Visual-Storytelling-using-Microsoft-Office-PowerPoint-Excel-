// Builds Day_1_PowerPoint_Exercises.pptx, Capstone_PowerPoint_Template.pptx and solutions.pptx (for trainer-guide images)
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const fs = require("fs");
const { applyTheme } = require("/root/.claude/skills/synced/25a79d21-886f-4f60-a430-4c2ac08b757f_bd3fa435-35ce-4a47-a092-5fe31c88cc92/pptx/scripts/apply_theme.js");

const S = JSON.parse(fs.readFileSync("/home/claude/work/stats.json"));
const REG = ["Central", "Northern", "Southern", "East Coast", "East Malaysia"];
const H = S.h1, Q = S.q3;
const pct = (x, d = 1) => (x * 100).toFixed(d) + "%";
const myr = (x) => "MYR " + Math.round(x).toLocaleString("en-US");
const num = (x) => Math.round(x).toLocaleString("en-US");

const THEME = {
  name: "Teratai Teal", headFontFace: "Arial", bodyFontFace: "Arial",
  colors: { dk1: "1F2937", lt1: "FFFFFF", dk2: "134E4A", lt2: "F0FDFA", accent1: "0F766E", accent2: "EA580C",
    accent3: "94A3B8", accent4: "14B8A6", accent5: "F59E0B", accent6: "64748B", hlink: "0F766E", folHlink: "64748B" },
};
const HEX = { teal: "0F766E", dark: "134E4A", light: "F0FDFA", orange: "EA580C", grey: "94A3B8", slate: "64748B", ink: "1F2937", mint: "CCFBF1", line: "E2E8F0", peach: "FFEDD5" };

async function icon(Comp, color = "#FFFFFF", size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

function baseLayouts(pres, footerText) {
  const C = pres.SchemeColor;
  const footer = [
    { text: { text: footerText, options: { x: 0.6, y: 7.02, w: 9, h: 0.3, fontSize: 10, color: C.accent6, margin: 0 } } },
  ];
  const slideNumber = { x: 12.2, y: 7.02, w: 0.6, h: 0.3, fontSize: 10, color: C.accent6, align: "right" };
  pres.defineSlideMaster({ title: "TITLE", background: { color: HEX.dark },
    objects: [{ placeholder: { options: { name: "title", type: "title", align: "left", x: 0.8, y: 2.2, w: 11.7, h: 1.6, fontSize: 40, bold: true, color: C.background1, valign: "bottom", margin: 0 }, text: "Title" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.0, w: 11.7, h: 1.6, fontSize: 18, color: "CCFBF1", valign: "top", margin: 0 }, text: "Subtitle" } }] });
  pres.defineSlideMaster({ title: "SECTION", background: { color: HEX.teal },
    objects: [{ placeholder: { options: { name: "body", type: "body", x: 0.8, y: 2.1, w: 11.7, h: 0.6, fontSize: 20, bold: true, color: "FDBA74", margin: 0 }, text: "Exercise" } },
      { placeholder: { options: { name: "title", type: "title", align: "left", x: 0.8, y: 2.8, w: 11.7, h: 1.4, fontSize: 40, bold: true, color: C.background1, valign: "top", margin: 0 }, text: "Section" } },
      { placeholder: { options: { name: "sub", type: "body", x: 0.8, y: 4.3, w: 11.7, h: 1.2, fontSize: 16, color: "CCFBF1", valign: "top", margin: 0 }, text: "Description" } }] });
  pres.defineSlideMaster({ title: "CONTENT", background: { color: "FFFFFF" }, slideNumber,
    objects: [...footer, { placeholder: { options: { name: "title", type: "title", align: "left", x: 0.6, y: 0.35, w: 12.1, h: 0.95, fontSize: 28, bold: true, color: C.text2, valign: "middle", margin: 0 }, text: "Title" } }] });
  return { slideNumber, footer };
}

function tagLayout(pres, name, tag, tagColor, footer, slideNumber, extra = []) {
  const C = pres.SchemeColor;
  pres.defineSlideMaster({ title: name, background: { color: "FFFFFF" }, slideNumber,
    objects: [...footer,
      { rect: { x: 10.55, y: 0.45, w: 2.2, h: 0.42, fill: { color: tagColor }, rectRadius: 0.1 } },
      { text: { text: tag, options: { x: 10.55, y: 0.45, w: 2.2, h: 0.42, fontSize: 11, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0 } } },
      { placeholder: { options: { name: "title", type: "title", align: "left", x: 0.6, y: 0.35, w: 9.7, h: 0.95, fontSize: 28, bold: true, color: C.text2, valign: "middle", margin: 0 }, text: "Click to add title" } },
      ...extra] });
}

// --- reusable card ---
function card(slide, pres, x, y, w, h, head, body, opts = {}) {
  const C = pres.SchemeColor;
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: opts.fill || HEX.light }, line: { color: opts.line || HEX.mint, width: 1 }, rectRadius: 0.08, objectName: "card " + head });
  slide.addText(head, { x: x + 0.2, y: y + 0.15, w: w - 0.4, h: 0.4, fontSize: 15, bold: true, color: opts.headColor || C.accent1, margin: 0, isTextBox: true });
  slide.addText(body, { x: x + 0.2, y: y + 0.6, w: w - 0.4, h: h - 0.75, fontSize: opts.size || 14, color: C.text1, valign: "top", margin: 0, isTextBox: true, paraSpaceAfter: 4 });
}
const bullets = (arr) => arr.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < arr.length - 1 } }));

function table(slide, rows, x, y, w, colW, fs = 12, highlightRow = -1) {
  const t = rows.map((r, i) => r.map((v, j) => ({ text: String(v), options: {
    bold: i === 0 || i === highlightRow, color: i === 0 ? "FFFFFF" : (i === highlightRow ? HEX.orange : HEX.ink),
    fill: { color: i === 0 ? HEX.teal : (i % 2 ? "FFFFFF" : "F8FAFC") }, align: j === 0 ? "left" : "right", fontSize: fs } })));
  slide.addTable(t, { x, y, w, colW, border: { type: "solid", color: HEX.line, pt: 0.75 }, rowH: 0.32, margin: [2, 6, 2, 6] });
}

// ---------------- data tables used in Day 1 ----------------
const regRows = REG.map((r) => [r, H.by_region[r].orders, H.by_region[r].value, H.by_region[r].value / H.total.value]);
const tableA = [["Region", "Orders", "Order Value (MYR)", "Share of value"], ...regRows.map((r) => [r[0], r[1], num(r[2]), pct(r[3])]),
  ["Total", H.total.orders, num(H.total.value), "100.0%"]];
const tableB = [["Month (2026)", "Orders", "Order Value (MYR)", "On-time rate"], ...H.months.map((m) => [m, H.by_month[m].orders, num(H.by_month[m].value), pct(H.by_month[m].rate)])];
const tableC = [["Region", "On time", "Late", "In transit", "On-time rate"], ...REG.map((r) => { const k = H.by_region[r]; return [r, k.on_time, k.late, k.in_transit, pct(k.rate)]; }),
  ["Total", H.total.on_time, H.total.late, H.total.in_transit, pct(H.total.rate)]];

// =====================================================================================
async function buildDay1() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.theme = { headFontFace: "Arial", bodyFontFace: "Arial" };
  pres.title = "Day 1 PowerPoint Exercises"; pres.author = "Fakhrul Syahmi";
  const C = pres.SchemeColor;
  const { slideNumber, footer } = baseLayouts(pres, "Teratai Distribution (fictional case) · Day 1 · PowerPoint exercises");
  tagLayout(pres, "BRIEF", "EXERCISE BRIEF", HEX.teal, footer, slideNumber);
  tagLayout(pres, "SOURCE", "SOURCE CONTENT", HEX.slate, footer, slideNumber);
  tagLayout(pres, "WORK", "WORKING SLIDE", HEX.orange, footer, slideNumber);
  tagLayout(pres, "WORKBODY", "WORKING SLIDE", HEX.orange, footer, slideNumber, [
    { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 1.55, w: 12.1, h: 5.2, fontSize: 16, color: C.text1, valign: "top" }, text: "Process steps" } }]);
  // raw cluttered slide layouts: only a small tag at the bottom
  for (const [nm, tg, col] of [["RAW", "SOURCE · original cluttered slide · do not edit", HEX.slate], ["RAWWORK", "WORKING SLIDE · redesign this copy", HEX.orange]]) {
    pres.defineSlideMaster({ title: nm, background: { color: "FFFFFF" },
      objects: [{ rect: { x: 0, y: 7.1, w: 4.6, h: 0.4, fill: { color: col } } },
        { text: { text: tg, options: { x: 0.15, y: 7.1, w: 4.4, h: 0.4, fontSize: 11, bold: true, color: "FFFFFF", valign: "middle", margin: 0 } } },
        { placeholder: { options: { name: "title", type: "title", align: "left", x: 0.3, y: 0.15, w: 12.7, h: 0.9, fontSize: 22, color: "C00000", fontFace: "Times New Roman", bold: true, align: "center", margin: 0 }, text: "Title" } }] });
  }

  // 1 Title
  pres.addSection({ title: "Introduction" });
  let s = pres.addSlide({ masterName: "TITLE", sectionTitle: "Introduction" });
  s.addText("Day 1 · PowerPoint Exercises", { placeholder: "title" });
  s.addText("Mastering Data Visualisation & Visual Storytelling using Microsoft Office (PowerPoint & Excel)\nTrainer: Fakhrul Syahmi  ·  Case: Teratai Distribution Sdn Bhd (fictional)", { placeholder: "body" });

  // 2 How to use
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Introduction" });
  s.addText("How to use this deck", { placeholder: "title" });
  const icoBrief = await icon(fa.FaBookOpen), icoSrc = await icon(fa.FaDatabase), icoWork = await icon(fa.FaPencilAlt);
  const legend = [[icoBrief, HEX.teal, "Exercise brief", "Scenario, task, source content and the expected output. Read first."],
    [icoSrc, HEX.slate, "Source content", "The material you start from. Leave it unchanged so you can compare before and after."],
    [icoWork, HEX.orange, "Working slide", "Your editable copy. Do the exercise here. Duplicate it (Ctrl+D) if you want to try two versions."]];
  legend.forEach(([ic, col, h, b], i) => {
    const x = 0.6 + i * 4.1;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.6, w: 3.8, h: 2.9, fill: { color: HEX.light }, line: { color: HEX.mint }, rectRadius: 0.08 });
    s.addShape(pres.shapes.OVAL, { x: x + 0.25, y: 1.85, w: 0.8, h: 0.8, fill: { color: col }, line: { color: col } });
    s.addImage({ data: ic, x: x + 0.45, y: 2.05, w: 0.4, h: 0.4, altText: h });
    s.addText(h, { x: x + 0.25, y: 2.8, w: 3.3, h: 0.45, fontSize: 18, bold: true, color: col, margin: 0, isTextBox: true });
    s.addText(b, { x: x + 0.25, y: 3.3, w: 3.3, h: 1.1, fontSize: 14, color: C.text1, margin: 0, valign: "top", isTextBox: true });
  });
  const flow = [["Exercise 1", "Slide redesign", "35 min"], ["Exercise 2", "Process infographic", "30 min"], ["Exercise 3", "Chart selection", "40 min"]];
  flow.forEach(([a, b, t], i) => {
    const x = 0.6 + i * 4.1;
    s.addShape(pres.shapes.CHEVRON, { x, y: 4.95, w: 3.8, h: 1.1, fill: { color: i === 0 ? HEX.teal : (i === 1 ? "14B8A6" : HEX.dark) }, line: { type: "none" } });
    s.addText([{ text: a + "  ·  " + t, options: { fontSize: 12, color: "CCFBF1", breakLine: true } }, { text: b, options: { fontSize: 18, bold: true, color: "FFFFFF" } }],
      { x: x + 0.5, y: 4.95, w: 2.9, h: 1.1, valign: "middle", margin: 0, isTextBox: true });
  });
  s.addText("Sample solutions are in the Trainer Answer Guide.", { x: 0.6, y: 6.3, w: 12, h: 0.4, fontSize: 12, italic: true, color: C.accent6, margin: 0, isTextBox: true });

  // 3 Case
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Introduction" });
  s.addText("The case: Teratai Distribution Sdn Bhd", { placeholder: "title" });
  card(s, pres, 0.6, 1.55, 5.9, 2.55, "Who they are", "A fictional FMCG distributor based in Shah Alam, Selangor. It delivers Beverages, Packaged Food, Personal Care and Household Care products to retailers in five regions.");
  card(s, pres, 0.6, 4.3, 5.9, 2.4, "What management monitors", bullets(["Monthly order volume and order value (MYR)", "On-time delivery rate by region", "Company target: 90% of completed orders on time"]));
  s.addText("Standard lead times and H1 2026 snapshot", { x: 6.9, y: 1.55, w: 5.8, h: 0.4, fontSize: 15, bold: true, color: C.accent1, margin: 0, isTextBox: true });
  const lt = { Central: 2, Northern: 3, Southern: 3, "East Coast": 4, "East Malaysia": 7 };
  table(s, [["Region", "Lead time (days)", "Orders Jan–Jun"], ...REG.map((r) => [r, lt[r], H.by_region[r].orders]), ["Total", "", H.total.orders]], 6.9, 2.05, 5.8, [2.4, 1.7, 1.7], 13);
  s.addText("On-time delivery = Actual Delivery Date on or before Promised Delivery Date. On-time rate = on-time deliveries ÷ completed orders (in-transit orders excluded).",
    { x: 6.9, y: 4.6, w: 5.8, h: 1.2, fontSize: 13, color: C.text1, margin: 0, valign: "top", isTextBox: true });

  // ---- Exercise 1 ----
  pres.addSection({ title: "Exercise 1 · Slide redesign" });
  s = pres.addSlide({ masterName: "SECTION", sectionTitle: "Exercise 1 · Slide redesign" });
  s.addText("EXERCISE 1", { placeholder: "body" }); s.addText("Slide redesign", { placeholder: "title" });
  s.addText("Turn a cluttered management slide into one clear message.", { placeholder: "sub" });

  s = pres.addSlide({ masterName: "BRIEF", sectionTitle: "Exercise 1 · Slide redesign" });
  s.addText("Exercise 1 brief: find the message, then design for it", { placeholder: "title" });
  card(s, pres, 0.6, 1.55, 5.95, 2.45, "Scenario", "The Operations Manager prepared this slide for the Monday management meeting. The Director said: “I can’t tell what you want me to notice.” You have 10 minutes to fix it before the meeting.");
  card(s, pres, 6.75, 1.55, 5.95, 2.45, "Source content", "Next slide: the original cluttered slide with a native chart (Orders, On-time % and Late deliveries by region, Jan–Jun 2026) and six competing text boxes.");
  card(s, pres, 0.6, 4.2, 5.95, 2.6, "Your task (on the working slide)", bullets(["Identify the ONE main finding", "Rewrite the headline as a full-sentence takeaway", "Improve hierarchy: remove, group or move competing messages", "Highlight the data that proves the finding (colour, label, callout)"]));
  card(s, pres, 6.75, 4.2, 5.95, 2.6, "Expected output", bullets(["One slide a director understands in 5 seconds", "Headline states the finding and the number", "Chart shows only the data that supports it; one accent colour", "Side notes (stock-take reminder etc.) removed or moved"]), { fill: HEX.peach, line: "FED7AA", headColor: HEX.orange });

  const clutter = (slide) => {
    const data = [
      { name: "Orders", labels: REG, values: REG.map((r) => H.by_region[r].orders) },
      { name: "On-Time %", labels: REG, values: REG.map((r) => +(H.by_region[r].rate * 100).toFixed(1)) },
      { name: "Late Deliveries", labels: REG, values: REG.map((r) => H.by_region[r].late) },
    ];
    slide.addText("Monthly Operations Update H1 2026 – Regional Delivery Performance, Orders & Other Updates!!", { placeholder: "title" });
    slide.addChart(pres.charts.BAR, data, { x: 0.3, y: 1.15, w: 7.6, h: 5.8, barDir: "col", barGrouping: "clustered",
      chartColors: ["FF0000", "00B050", "7030A0"], showValue: true, dataLabelFontSize: 9, dataLabelColor: "000000", showLegend: true, legendPos: "t", legendFontSize: 9,
      showTitle: true, title: "Orders, On-Time % and Late Deliveries by Region (Jan-Jun 2026) - all regions", titleFontSize: 11, titleColor: "000000",
      valGridLine: { color: "808080", size: 1 }, catGridLine: { color: "BFBFBF", size: 1 }, catAxisLabelFontSize: 9, valAxisLabelFontSize: 9,
      plotArea: { fill: { color: "FFF2CC" } }, objectName: "cluttered chart" });
    const boxes = [
      ["Central = our BIGGEST region!! " + H.by_region.Central.orders + " orders", 8.1, 1.2, 4.9, 0.6, "0070C0", 16, "DDEBF7"],
      ["Total order value " + myr(H.total.value) + " (Jan–Jun)", 8.1, 1.9, 4.9, 0.55, "00B050", 14, "E2EFDA"],
      ["East Malaysia: " + H.by_region["East Malaysia"].late + " late deliveries?? check with 3PL", 8.1, 2.55, 4.9, 0.6, "C00000", 12, "FFFFFF"],
      ["Packaged Food = top category by value (" + myr(H.by_cat["Packaged Food"].value) + ")", 8.1, 3.25, 4.9, 0.6, "7030A0", 13, "EDE7F6"],
      ["Target on-time = 90%", 8.1, 3.95, 2.3, 0.75, "000000", 14, "FFFF00"],
      ["REMINDER: Stock-take at Shah Alam DC on 15 July – all staff!", 10.5, 3.95, 2.5, 0.95, "FF0000", 11, "FCE4D6"],
      ["Overall on-time rate " + pct(H.total.rate) + " (" + H.total.on_time + "/" + H.total.completed + " completed). " + H.total.in_transit + " orders still in transit at 30 Jun. Southern best at " + pct(H.by_region.Southern.rate) + ". New hire in planning team starts Aug.", 8.1, 5.05, 4.9, 0.95, "595959", 10, "F2F2F2"],
    ];
    boxes.forEach(([t, x, y, w, h, col, fsz, fill], i) => slide.addText(t, { x, y, w, h, fontSize: fsz, color: col, bold: i < 2, fontFace: i % 2 ? "Comic Sans MS" : "Arial", fill: { color: fill }, line: { color: col, width: 1 }, valign: "middle", isTextBox: true, objectName: "clutter box " + (i + 1) }));
    slide.addShape(pres.shapes.STAR_16_POINT, { x: 11.8, y: 6.05, w: 1.1, h: 0.95, fill: { color: "FFC000" }, line: { color: "FF0000" }, objectName: "NEW star" });
    slide.addText("NEW!", { x: 11.8, y: 6.05, w: 1.1, h: 0.95, fontSize: 12, bold: true, color: "C00000", align: "center", valign: "middle", isTextBox: true });
    slide.addText("Source: order system, extracted 30 Jun 2026. Figures unaudited. Prepared by Ops.", { x: 8.1, y: 6.2, w: 3.6, h: 0.5, fontSize: 8, color: "7F7F7F", isTextBox: true });
  };
  s = pres.addSlide({ masterName: "RAW", sectionTitle: "Exercise 1 · Slide redesign" }); clutter(s);
  s.addNotes("Source slide for Exercise 1. Do not edit. Use the next slide (working copy).");
  s = pres.addSlide({ masterName: "RAWWORK", sectionTitle: "Exercise 1 · Slide redesign" }); clutter(s);
  s.addNotes("Working copy. Steps: 1) Decide the main finding. 2) Replace the title with a takeaway headline. 3) Delete or move competing boxes. 4) Right-click the chart > Edit Data: keep only On-Time %, sort it, colour East Malaysia in the accent colour and the rest grey. 5) Add a callout with the supporting numbers.");

  // ---- Exercise 2 ----
  pres.addSection({ title: "Exercise 2 · Infographic" });
  s = pres.addSlide({ masterName: "SECTION", sectionTitle: "Exercise 2 · Infographic" });
  s.addText("EXERCISE 2", { placeholder: "body" }); s.addText("Process infographic", { placeholder: "title" });
  s.addText("Turn a text-heavy process description into a visual flow.", { placeholder: "sub" });
  s = pres.addSlide({ masterName: "BRIEF", sectionTitle: "Exercise 2 · Infographic" });
  s.addText("Exercise 2 brief: show the order fulfilment process visually", { placeholder: "title" });
  card(s, pres, 0.6, 1.55, 5.95, 2.45, "Scenario", "New customer-service staff keep asking where an order is “stuck”. The Operations Manager wants a one-slide visual of the order fulfilment process for the onboarding pack and the DC notice board.");
  card(s, pres, 6.75, 1.55, 5.95, 2.45, "Source content", "Next slide: the six process steps in plain text, with the owner and the typical time for each step.");
  card(s, pres, 0.6, 4.2, 5.95, 2.6, "Your task (on the working slide)", bullets(["Click the text box > Home > Convert to SmartArt > choose a Process layout, OR build it with shapes (chevrons, circles, arrows)", "Add one icon per step (Insert > Icons)", "Keep step names short; move details to small labels"]));
  card(s, pres, 6.75, 4.2, 5.95, 2.6, "Expected output", bullets(["Six steps read left-to-right in under 10 seconds", "Consistent shape, icon and colour for each step", "Owner and time visible but secondary", "Lead-time rule shown once, not repeated"]), { fill: HEX.peach, line: "FED7AA", headColor: HEX.orange });

  const STEPS = [
    ["Order received", "Sales rep or B2B portal captures the retailer's order", "Sales Admin", "Same day"],
    ["Credit & stock check", "Finance checks credit limit; system reserves stock", "Finance / Inventory", "Within 4 hours"],
    ["Pick & pack", "Shah Alam DC picks, packs and labels the order", "DC Warehouse team", "Within 1 day"],
    ["Dispatch & route planning", "Transport planner assigns the truck or 3PL and the route", "Transport Planning", "Same day as packing"],
    ["Delivery to customer", "Driver delivers; East Malaysia orders go by sea freight + local 3PL", "Driver / 3PL partner", "Per regional lead time"],
    ["Proof of delivery & invoice", "Customer signs e-POD; invoice issued automatically", "Driver / Billing", "Same day as delivery"],
  ];
  const stepText = STEPS.map((st, i) => `Step ${i + 1}: ${st[0]}. ${st[1]}. Owner: ${st[2]}. Time: ${st[3]}.`);
  s = pres.addSlide({ masterName: "SOURCE", sectionTitle: "Exercise 2 · Infographic" });
  s.addText("Order fulfilment process, as written by Operations", { placeholder: "title" });
  s.addText(stepText.map((t, i) => ({ text: t, options: { breakLine: i < stepText.length - 1 } })),
    { x: 0.6, y: 1.5, w: 12.1, h: 4.3, fontSize: 15, color: HEX.ink, paraSpaceAfter: 8, valign: "top", isTextBox: true });
  s.addText("Lead-time rule: Promised Delivery Date = Order Date + 2 days (Central), 3 days (Northern, Southern), 4 days (East Coast), 7 days (East Malaysia).",
    { x: 0.6, y: 5.9, w: 12.1, h: 0.7, fontSize: 14, italic: true, color: C.accent6, isTextBox: true });
  s = pres.addSlide({ masterName: "WORKBODY", sectionTitle: "Exercise 2 · Infographic" });
  s.addText("From order to invoice in 6 steps", { placeholder: "title" });
  s.addText(STEPS.map((st, i) => ({ text: st[0] + " — " + st[2] + ", " + st[3], options: { bullet: true, breakLine: i < STEPS.length - 1 } })), { placeholder: "body" });
  s.addNotes("Click inside the bulleted text, then Home > Convert to SmartArt > More SmartArt Graphics > Process (e.g. Basic Chevron Process or Step Up Process). Then add icons with Insert > Icons and recolour with SmartArt Design > Change Colours. Alternative: delete the text box and build the flow with Insert > Shapes.");

  // ---- Exercise 3 ----
  pres.addSection({ title: "Exercise 3 · Chart selection" });
  s = pres.addSlide({ masterName: "SECTION", sectionTitle: "Exercise 3 · Chart selection" });
  s.addText("EXERCISE 3", { placeholder: "body" }); s.addText("Chart selection", { placeholder: "title" });
  s.addText("Match each business question to the right chart and write the takeaway.", { placeholder: "sub" });
  s = pres.addSlide({ masterName: "BRIEF", sectionTitle: "Exercise 3 · Chart selection" });
  s.addText("Exercise 3 brief: the question decides the chart", { placeholder: "title" });
  card(s, pres, 0.6, 1.55, 5.95, 2.45, "Scenario", "The Director will ask three questions at the half-year review. You have the numbers but no charts. Each answer needs one chart and one sentence.");
  card(s, pres, 6.75, 1.55, 5.95, 2.45, "Source content", "Next slide: three compact tables (regional orders, monthly performance, delivery status by region) and the three business questions.");
  card(s, pres, 0.6, 4.2, 5.95, 2.6, "Your task (working slides A, B, C)", bullets(["Pick the chart type that answers the question (comparison, trend, part-to-whole or ranking)", "Insert > Chart and paste only the data you need", "Write a one-sentence takeaway as the slide headline", "Note why you chose that chart"]));
  card(s, pres, 6.75, 4.2, 5.95, 2.6, "Expected output", bullets(["Three native, editable charts", "Each headline answers its question with a number", "Sorted where ranking matters; one highlight colour", "No 3-D, no pie with more than 3 slices"]), { fill: HEX.peach, line: "FED7AA", headColor: HEX.orange });

  s = pres.addSlide({ masterName: "SOURCE", sectionTitle: "Exercise 3 · Chart selection" });
  s.addText("Source tables: Teratai Distribution, January–June 2026", { placeholder: "title" });
  s.addText("A · Regional orders", { x: 0.6, y: 1.4, w: 6, h: 0.35, fontSize: 14, bold: true, color: C.accent1, margin: 0, isTextBox: true });
  table(s, tableA, 0.6, 1.8, 6.0, [1.7, 0.9, 1.9, 1.5], 11);
  s.addText("B · Monthly performance", { x: 6.9, y: 1.4, w: 6, h: 0.35, fontSize: 14, bold: true, color: C.accent1, margin: 0, isTextBox: true });
  table(s, tableB, 6.9, 1.8, 5.8, [1.5, 0.9, 1.9, 1.5], 11);
  s.addText("C · Delivery status by region (on-time rate = on time ÷ completed)", { x: 0.6, y: 4.3, w: 6.2, h: 0.35, fontSize: 14, bold: true, color: C.accent1, margin: 0, isTextBox: true });
  table(s, tableC, 0.6, 4.7, 6.0, [1.7, 1.0, 0.8, 1.1, 1.4], 11);
  card(s, pres, 6.9, 4.3, 5.8, 2.45, "Business questions", [
    { text: "A  Which region contributes the most order value, and by how much?", options: { breakLine: true } },
    { text: "B  How did monthly order volume change from January to June?", options: { breakLine: true } },
    { text: "C  Which region has the weakest on-time delivery, compared with the 90% target?" }], { size: 13 });

  const work3 = [
    ["A", "Which region contributes the most order value, and by how much?", tableA],
    ["B", "How did monthly order volume change from January to June?", tableB],
    ["C", "Which region has the weakest on-time delivery, compared with the 90% target?", tableC],
  ];
  for (const [k, q, tb] of work3) {
    s = pres.addSlide({ masterName: "WORK", sectionTitle: "Exercise 3 · Chart selection" });
    s.addText("Takeaway " + k + ": one sentence that answers the question", { placeholder: "title" });
    s.addText([{ text: "Question " + k + "  ", options: { bold: true, color: HEX.orange } }, { text: q }], { x: 0.6, y: 1.35, w: 12.1, h: 0.45, fontSize: 15, color: HEX.ink, margin: 0, isTextBox: true });
    s.addText("Source data (copy what you need)", { x: 0.6, y: 1.95, w: 4.7, h: 0.3, fontSize: 12, bold: true, color: C.accent6, margin: 0, isTextBox: true });
    const cw = tb[0].length === 5 ? [1.4, 0.75, 0.6, 0.85, 1.0] : [1.4, 0.8, 1.45, 0.95];
    table(s, tb, 0.6, 2.3, cw.reduce((a, b) => a + b), cw, 10);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.6, y: 1.95, w: 7.1, h: 4.0, fill: { color: "FFFFFF" }, line: { color: HEX.teal, width: 1.25, dashType: "dash" }, rectRadius: 0.06, objectName: "chart placeholder" });
    s.addText("Insert ▸ Chart here\n(delete this frame when your chart is in place)", { x: 5.6, y: 3.45, w: 7.1, h: 1.0, fontSize: 14, color: C.accent1, align: "center", isTextBox: true, objectName: "chart placeholder text" });
    s.addText([{ text: "Chart chosen: ", options: { bold: true } }, { text: "__________   " }, { text: "Why: ", options: { bold: true } }, { text: "______________________________________" }],
      { x: 5.6, y: 6.1, w: 7.1, h: 0.55, fontSize: 13, color: HEX.ink, fill: { color: HEX.peach }, margin: 6, isTextBox: true, objectName: "chart rationale" });
  }

  // wrap-up
  pres.addSection({ title: "Wrap-up" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Wrap-up" });
  s.addText("Before you finish: five-point slide check", { placeholder: "title" });
  const checks = [[fa.FaBullseye, "One message", "Can someone say the main point after 5 seconds?"], [fa.FaHeading, "Headline = takeaway", "Full sentence with the key number, not a topic label."],
    [fa.FaPalette, "One accent colour", "Grey for context, colour only for the data that matters."], [fa.FaChartBar, "Right chart", "Ranking = sorted bar, trend = line, part-to-whole = stacked bar."],
    [fa.FaEraser, "Less is more", "Remove gridlines, legends and boxes that do not support the message."]];
  for (let i = 0; i < checks.length; i++) {
    const [ic, h, b] = checks[i];
    const y = 1.5 + i * 1.05;
    s.addShape(pres.shapes.OVAL, { x: 0.7, y, w: 0.75, h: 0.75, fill: { color: i % 2 ? HEX.orange : HEX.teal }, line: { type: "none" } });
    s.addImage({ data: await icon(ic), x: 0.89, y: y + 0.19, w: 0.37, h: 0.37, altText: h });
    s.addText(h, { x: 1.7, y, w: 3.3, h: 0.75, fontSize: 18, bold: true, color: C.text2, valign: "middle", margin: 0, isTextBox: true });
    s.addText(b, { x: 5.0, y, w: 7.6, h: 0.75, fontSize: 15, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
  }
  await pres.writeFile({ fileName: "/home/claude/work/out/Day_1_PowerPoint_Exercises.pptx" });
  await applyTheme("/home/claude/work/out/Day_1_PowerPoint_Exercises.pptx", THEME);
  return STEPS;
}

// =====================================================================================
async function buildSolutions(STEPS) {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.theme = { headFontFace: "Arial", bodyFontFace: "Arial" };
  const C = pres.SchemeColor;
  const { slideNumber, footer } = baseLayouts(pres, "Sample solution · Trainer Answer Guide");
  const axis = { catAxisLabelColor: "475569", valAxisLabelColor: "475569", catAxisLabelFontSize: 12, valAxisLabelFontSize: 11, catGridLine: { style: "none" }, valGridLine: { color: "E2E8F0", size: 0.75 }, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt" };

  // S1 redesign
  const sorted = [...REG].sort((a, b) => H.by_region[b].rate - H.by_region[a].rate);
  let s = pres.addSlide({ masterName: "CONTENT" });
  s.addText("East Malaysia delivers only " + pct(H.by_region["East Malaysia"].rate, 0) + " of orders on time — well below the 90% target", { placeholder: "title" });
  s.addChart(pres.charts.BAR, [{ name: "On-time rate", labels: sorted, values: sorted.map((r) => H.by_region[r].rate) }], {
    x: 0.6, y: 1.5, w: 7.6, h: 5.2, barDir: "bar", catAxisOrientation: "maxMin", valAxisMinVal: 0, valAxisMaxVal: 1, valAxisLabelFormatCode: "0%", valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    chartColors: sorted.map((r) => (r === "East Malaysia" ? HEX.orange : "CBD5E1")), showValue: true, dataLabelFormatCode: "0.0%", dataLabelPosition: "outEnd", dataLabelFontSize: 13, dataLabelColor: "1F2937",
    showLegend: false, showTitle: true, title: "On-time delivery rate, completed orders, Jan–Jun 2026", titleFontSize: 13, titleColor: "475569", catAxisLabelColor: "1F2937", catAxisLabelFontSize: 13, barGapWidthPct: 60, catAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt" });
  const facts = [[pct(H.by_region["East Malaysia"].rate), "East Malaysia on time (" + H.by_region["East Malaysia"].on_time + " of " + H.by_region["East Malaysia"].completed + ")"],
    [pct(S.h1_others.rate), "All other regions combined"], [pct(S.h1_em_late_share, 0), "of late deliveries are East Malaysia, which has " + pct(S.h1_em_order_share, 0) + " of orders"]];
  facts.forEach(([n, l], i) => {
    s.addText(n, { x: 8.6, y: 1.6 + i * 1.7, w: 4.1, h: 0.8, fontSize: 40, bold: true, color: i === 1 ? HEX.slate : HEX.orange, margin: 0, isTextBox: true });
    s.addText(l, { x: 8.6, y: 2.4 + i * 1.7, w: 4.1, h: 0.7, fontSize: 14, color: HEX.ink, margin: 0, valign: "top", isTextBox: true });
  });

  // S2 process
  s = pres.addSlide({ masterName: "CONTENT" });
  s.addText("How an order reaches the customer: 6 steps from order to invoice", { placeholder: "title" });
  const icons = [fa.FaShoppingCart, fa.FaClipboardCheck, fa.FaBoxOpen, fa.FaRoute, fa.FaTruck, fa.FaFileInvoiceDollar];
  for (let i = 0; i < 6; i++) {
    const x = 0.6 + i * 2.05;
    s.addShape(pres.shapes.CHEVRON, { x, y: 1.9, w: 2.1, h: 1.0, fill: { color: i === 4 ? HEX.orange : HEX.teal }, line: { color: "FFFFFF", width: 1 } });
    s.addText(String(i + 1), { x: x + 0.3, y: 1.9, w: 1.5, h: 1.0, fontSize: 24, bold: true, color: "FFFFFF", align: "center", valign: "middle", isTextBox: true });
    s.addShape(pres.shapes.OVAL, { x: x + 0.55, y: 3.15, w: 0.95, h: 0.95, fill: { color: HEX.light }, line: { color: i === 4 ? HEX.orange : HEX.teal, width: 1.5 } });
    s.addImage({ data: await icon(icons[i], "#" + (i === 4 ? HEX.orange : HEX.teal)), x: x + 0.79, y: 3.39, w: 0.47, h: 0.47 });
    s.addText(STEPS[i][0], { x: x + 0.05, y: 4.25, w: 1.95, h: 0.75, fontSize: 14, bold: true, color: HEX.dark, align: "center", valign: "top", margin: 0, isTextBox: true });
    s.addText(STEPS[i][2] + "\n" + STEPS[i][3], { x: x + 0.05, y: 5.0, w: 1.95, h: 0.8, fontSize: 11, color: HEX.slate, align: "center", valign: "top", margin: 0, isTextBox: true });
  }
  s.addText("Promised delivery = order date + 2 days Central · 3 Northern & Southern · 4 East Coast · 7 East Malaysia", { x: 0.6, y: 6.15, w: 12.1, h: 0.5, fontSize: 13, italic: true, color: HEX.slate, align: "center", fill: { color: HEX.light }, isTextBox: true });

  // S3 chart selection (3 charts)
  s = pres.addSlide({ masterName: "CONTENT" });
  s.addText("Exercise 3 sample answers: one chart per question", { placeholder: "title" });
  const byVal = [...REG].sort((a, b) => H.by_region[b].value - H.by_region[a].value);
  const tk = [
    ["A · Sorted bar (ranking)", "Central brings in " + pct(H.by_region.Central.value / H.total.value, 0) + " of order value (MYR " + (H.by_region.Central.value / 1e3).toFixed(0) + "k), about 1.5× Southern."],
    ["B · Line (trend)", "Monthly orders rose " + pct(H.by_month.Jun.orders / H.by_month.Jan.orders - 1, 0) + ", from " + H.by_month.Jan.orders + " in Jan to " + H.by_month.Jun.orders + " in Jun."],
    ["C · Sorted bar + target", "Only Southern meets the 90% target; East Malaysia is lowest at " + pct(H.by_region["East Malaysia"].rate) + "."],
  ];
  const mk = (i) => ({ x: 0.5 + i * 4.2, y: 2.45, w: 4.0, h: 4.25 });
  s.addChart(pres.charts.BAR, [{ name: "Order value (MYR k)", labels: byVal, values: byVal.map((r) => Math.round(H.by_region[r].value / 1000)) }], { ...mk(0), ...axis, barDir: "bar", catAxisOrientation: "maxMin", valAxisHidden: true, valGridLine: { style: "none" },
    chartColors: byVal.map((r) => (r === "Central" ? HEX.teal : "CBD5E1")), showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 11, showLegend: false, catAxisLabelFontSize: 11 });
  s.addChart(pres.charts.LINE, [{ name: "Orders", labels: H.months, values: H.months.map((m) => H.by_month[m].orders) }], { ...mk(1), ...axis, chartColors: [HEX.teal], lineSize: 3, lineDataSymbolSize: 8, showValue: true, dataLabelPosition: "t", dataLabelFontSize: 11, showLegend: false, valAxisMinVal: 0, valAxisMaxVal: 70 });
  const byRate = [...REG].sort((a, b) => H.by_region[b].rate - H.by_region[a].rate);
  s.addChart(pres.charts.BAR, [{ name: "On-time rate", labels: byRate, values: byRate.map((r) => H.by_region[r].rate) }], { ...mk(2), ...axis, barDir: "bar", catAxisOrientation: "maxMin", valAxisMinVal: 0, valAxisMaxVal: 1, valAxisHidden: true, valGridLine: { style: "none" },
    chartColors: byRate.map((r) => (r === "East Malaysia" ? HEX.orange : "CBD5E1")), showValue: true, dataLabelFormatCode: "0%", dataLabelPosition: "outEnd", dataLabelFontSize: 11, showLegend: false, catAxisLabelFontSize: 11 });
  tk.forEach(([h, t], i) => {
    s.addText(h, { x: 0.5 + i * 4.2, y: 1.4, w: 4.0, h: 0.35, fontSize: 14, bold: true, color: i === 2 ? HEX.orange : HEX.teal, margin: 0, isTextBox: true });
    s.addText(t, { x: 0.5 + i * 4.2, y: 1.75, w: 4.0, h: 0.7, fontSize: 12, color: HEX.ink, margin: 0, valign: "top", isTextBox: true });
  });

  // S4 capstone worked example insight
  s = pres.addSlide({ masterName: "CONTENT" });
  const emM = ["Jul", "Aug", "Sep"].map((m) => Q.region_month["East Malaysia"][m].rate);
  s.addText("East Malaysia on-time delivery fell from " + pct(emM[0], 0) + " in July to " + pct(emM[2], 0) + " in September, while other regions held at ~91%", { placeholder: "title" });
  s.addChart(pres.charts.LINE, [
    { name: "East Malaysia", labels: ["Jul", "Aug", "Sep"], values: emM },
    { name: "All other regions", labels: ["Jul", "Aug", "Sep"], values: ["Jul", "Aug", "Sep"].map((m) => S.q3_others_by_month[m].rate) },
  ], { x: 0.6, y: 1.55, w: 7.8, h: 5.1, ...axis, chartColors: [HEX.orange, "94A3B8"], lineSize: 3, lineDataSymbolSize: 9, valAxisMinVal: 0, valAxisMaxVal: 1, valAxisLabelFormatCode: "0%", showValue: true, dataLabelFormatCode: "0%", dataLabelPosition: "t", dataLabelFontSize: 12,
    showLegend: true, legendPos: "b", legendFontSize: 12, showTitle: true, title: "On-time delivery rate, completed orders, Q3 2026", titleFontSize: 13, titleColor: "475569" });
  const em = Q.by_region["East Malaysia"];
  const f2 = [[pct(em.rate), "East Malaysia Q3 on time (" + em.on_time + " of " + em.completed + " completed)"], [pct(S.q3_em_late_share, 0), "of all late deliveries, from " + pct(S.q3_em_order_share, 0) + " of orders"], [em.orders + " orders", "East Malaysia volume, up from " + Q.region_month["East Malaysia"].Jul.orders + " in Jul to " + Q.region_month["East Malaysia"].Sep.orders + " in Sep"]];
  f2.forEach(([n, l], i) => {
    s.addText(n, { x: 8.8, y: 1.6 + i * 1.7, w: 3.9, h: 0.8, fontSize: 36, bold: true, color: HEX.orange, margin: 0, isTextBox: true });
    s.addText(l, { x: 8.8, y: 2.4 + i * 1.7, w: 3.9, h: 0.7, fontSize: 14, color: HEX.ink, margin: 0, valign: "top", isTextBox: true });
  });
  await pres.writeFile({ fileName: "/home/claude/work/sol/solutions.pptx" });
  await applyTheme("/home/claude/work/sol/solutions.pptx", THEME);
}

// =====================================================================================
const PR = (t) => t.split("\n").map((x, i, a) => ({ text: x, options: { italic: true, color: "64748B", breakLine: i < a.length - 1 } }));
async function buildCapstone() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.theme = { headFontFace: "Arial", bodyFontFace: "Arial" };
  pres.title = "Capstone PowerPoint Template"; pres.author = "Fakhrul Syahmi";
  const C = pres.SchemeColor;
  const { slideNumber, footer } = baseLayouts(pres, "Teratai Distribution (fictional case) · Capstone · Q3 2026 delivery performance");
  const ph = (name, x, y, w, h, prompt, o = {}) => ({ placeholder: { options: { name, type: "body", x, y, w, h, fontSize: o.fs || 16, color: C.text1, valign: "top", ...o.extra }, text: prompt } });
  const lab = (t, x, y, w, col = HEX.teal) => ({ text: { text: t, options: { x, y, w, h: 0.4, fontSize: 14, bold: true, color: col, margin: 0 } } });
  const frame = (x, y, w, h, fill = HEX.light, line = HEX.mint) => ({ rect: { x, y, w, h, fill: { color: fill }, line: { color: line, width: 1 } } });

  pres.defineSlideMaster({ title: "CONTEXT", background: { color: "FFFFFF" }, slideNumber, objects: [...footer,
    { text: { text: "1 · CONTEXT", options: { x: 0.6, y: 0.3, w: 3, h: 0.35, fontSize: 12, bold: true, color: C.accent2, margin: 0 } } },
    { placeholder: { options: { name: "title", type: "title", align: "left", x: 0.6, y: 0.65, w: 12.1, h: 0.9, fontSize: 28, bold: true, color: C.text2, valign: "top", margin: 0 }, text: "Write the decision management must make (one sentence)" } },
    frame(0.6, 1.8, 5.9, 4.9), lab("What is management trying to understand or decide?", 0.85, 1.95, 5.5),
    ph("decision", 0.85, 2.45, 5.4, 4.1, "• Who is asking, and why now?\n• What decision will this analysis support?\n• What happens if nothing is done?"),
    frame(6.8, 1.8, 5.9, 4.9), lab("Scope of the analysis", 7.05, 1.95, 5.5),
    ph("scope", 7.05, 2.45, 5.4, 4.1, "• Data: Orders Q3 (Jul–Sep 2026), number of orders\n• KPIs used and how on-time rate is defined\n• Business questions you answered") ] });
  pres.defineSlideMaster({ title: "INSIGHT", background: { color: "FFFFFF" }, slideNumber, objects: [...footer,
    { text: { text: "2 · INSIGHT", options: { x: 0.6, y: 0.3, w: 3, h: 0.35, fontSize: 12, bold: true, color: C.accent2, margin: 0 } } },
    { placeholder: { options: { name: "title", type: "title", align: "left", x: 0.6, y: 0.65, w: 12.1, h: 0.9, fontSize: 28, bold: true, color: C.text2, valign: "top", margin: 0 }, text: "Takeaway headline: state the finding with its key number" } },
    frame(8.6, 1.8, 4.1, 4.9), lab("What the dashboard shows", 8.85, 1.95, 3.7),
    ph("evidence", 8.85, 2.45, 3.6, 4.1, "• Key number 1 (with order count)\n• Key number 2 (comparison)\n• Trend or pattern\n• One caveat about the data", { fs: 15 }) ] });
  pres.defineSlideMaster({ title: "RECO", background: { color: "FFFFFF" }, slideNumber, objects: [...footer,
    { text: { text: "3 · RECOMMENDATION", options: { x: 0.6, y: 0.3, w: 4, h: 0.35, fontSize: 12, bold: true, color: C.accent2, margin: 0 } } },
    { placeholder: { options: { name: "title", type: "title", align: "left", x: 0.6, y: 0.65, w: 12.1, h: 0.9, fontSize: 28, bold: true, color: C.text2, valign: "top", margin: 0 }, text: "Recommendation: the one action management should take" } },
    frame(0.6, 1.8, 3.85, 4.9, HEX.peach, "FED7AA"), lab("The action", 0.85, 1.95, 3.4, HEX.orange),
    ph("action", 0.85, 2.45, 3.35, 4.1, "What exactly should be done, by whom, and by when?"),
    frame(4.73, 1.8, 3.85, 4.9), lab("Why: link to the evidence", 4.98, 1.95, 3.4),
    ph("why", 4.98, 2.45, 3.35, 4.1, "Which finding on the Insight slide makes this the right action?"),
    frame(8.85, 1.8, 3.85, 4.9), lab("How we will know it worked", 9.1, 1.95, 3.4),
    ph("measure", 9.1, 2.45, 3.35, 4.1, "KPI to track, target, owner and review date. What still needs investigating?") ] });

  let s = pres.addSlide({ masterName: "TITLE" });
  s.addText("Delivery Performance Review · Q3 2026", { placeholder: "title" });
  s.addText("Group: ________   ·   Members: ____________________\nPrepared for: Puan Rosnah Idris, Operations Director, Teratai Distribution Sdn Bhd (fictional)", { placeholder: "body" });
  s.addNotes("Replace the group name and members. Keep this cover; it does not count toward the three content slides.");

  s = pres.addSlide({ masterName: "CONTEXT" });
  s.addText("Write the decision management must make (one sentence)", { placeholder: "title" });
  s.addText(PR("• Who is asking, and why now?\n• What decision will this analysis support?\n• What happens if nothing is done?"), { placeholder: "decision" });
  s.addText(PR("• Data: Orders Q3 (Jul–Sep 2026), number of orders\n• KPIs used and how on-time rate is defined\n• Business questions you answered"), { placeholder: "scope" });
  s.addNotes("Prompt: What is management trying to understand or decide? Keep to 60 seconds. Click each box and type over the prompt text.");
  s = pres.addSlide({ masterName: "INSIGHT" });
  s.addText("Takeaway headline: state the finding with its key number", { placeholder: "title" });
  s.addText(PR("• Key number 1 (with order count)\n• Key number 2 (comparison)\n• Trend or pattern\n• One caveat about the data"), { placeholder: "evidence" });
  s.addChart(pres.charts.BAR, [{ name: "Replace with your data", labels: ["Category 1", "Category 2", "Category 3", "Category 4", "Category 5"], values: [0.9, 0.85, 0.8, 0.75, 0.6] }], {
    x: 0.6, y: 1.8, w: 7.7, h: 4.9, barDir: "bar", catAxisOrientation: "maxMin", valAxisMinVal: 0, valAxisMaxVal: 1, valAxisLabelFormatCode: "0%", valGridLine: { color: "E2E8F0", size: 0.75 }, catGridLine: { style: "none" },
    chartColors: ["CBD5E1", "CBD5E1", "CBD5E1", "CBD5E1", HEX.orange], showValue: true, dataLabelFormatCode: "0%", dataLabelPosition: "outEnd", dataLabelFontSize: 12, showLegend: false,
    showTitle: true, title: "Chart title: metric, population, period (right-click ▸ Edit Data to replace)", titleFontSize: 12, titleColor: "475569", catAxisLabelColor: "475569", valAxisLabelColor: "475569",
    catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", objectName: "insight chart (editable)" });
  s.addNotes("Replace the sample chart: right-click > Edit Data, or delete it and paste your PivotChart from the Mini Dashboard (Paste Special > Use Destination Theme & Embed Workbook). Highlight only the bar or line that proves your headline.");
  s = pres.addSlide({ masterName: "RECO" });
  s.addText("Recommendation: the one action management should take", { placeholder: "title" });
  s.addText(PR("What exactly should be done, by whom, and by when?"), { placeholder: "action" });
  s.addText(PR("Which finding on the Insight slide makes this the right action?"), { placeholder: "why" });
  s.addText(PR("KPI to track, target, owner and review date. What still needs investigating?"), { placeholder: "measure" });
  s.addNotes("One action only. It must follow from the Insight slide. Include a measurable KPI (e.g. on-time rate target) and who owns it.");

  s = pres.addSlide({ masterName: "CONTENT" });
  s.addText("Presentation & peer feedback checklist", { placeholder: "title" });
  const hdr = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: HEX.teal }, fontSize: 13 } });
  const cell = (t, b) => ({ text: t, options: { fontSize: 12, bold: !!b, color: HEX.ink } });
  const box = { text: "☐", options: { fontSize: 16, align: "center", color: HEX.teal } };
  const sc = { text: "1  2  3  4", options: { fontSize: 12, align: "center", color: HEX.slate } };
  const rows = [[hdr("Area"), hdr("Check"), hdr("✓"), hdr("Score"), hdr("Peer comment")],
    [cell("Data accuracy", 1), cell("KPIs match the data; on-time rate uses completed orders only; order counts shown next to rates"), box, sc, cell("")],
    [cell("", 1), cell("Period, units (MYR, %) and source stated on the chart or slide"), box, sc, cell("")],
    [cell("Visual clarity", 1), cell("Headline states the finding; chart type fits the question; one highlight colour"), box, sc, cell("")],
    [cell("", 1), cell("Clutter removed: no 3-D, unnecessary gridlines, legends or decoration"), box, sc, cell("")],
    [cell("Recommendation", 1), cell("Action follows directly from the evidence on the Insight slide"), box, sc, cell("")],
    [cell("", 1), cell("Action is specific (owner, timing) and measurable (KPI, target); limits of the data acknowledged"), box, sc, cell("")],
    [cell("Delivery", 1), cell("Finished in 5 minutes; every member can explain the numbers"), box, sc, cell("")]];
  s.addTable(rows, { x: 0.6, y: 1.5, w: 12.1, colW: [1.9, 6.1, 0.6, 1.4, 2.1], border: { type: "solid", color: HEX.line, pt: 0.75 }, rowH: 0.52, valign: "middle", margin: [3, 6, 3, 6] });
  s.addText([{ text: "Peer feedback: ", options: { bold: true, color: HEX.orange } }, { text: "One thing that worked well: ____________________     One thing to improve: ____________________" }],
    { x: 0.6, y: 6.15, w: 12.1, h: 0.55, fontSize: 13, color: HEX.ink, fill: { color: HEX.peach }, margin: 6, isTextBox: true });
  await pres.writeFile({ fileName: "/home/claude/work/out/Capstone_PowerPoint_Template.pptx" });
  await applyTheme("/home/claude/work/out/Capstone_PowerPoint_Template.pptx", THEME);
}

(async () => {
  fs.mkdirSync("/home/claude/work/sol", { recursive: true });
  const STEPS = await buildDay1();
  await buildSolutions(STEPS);
  await buildCapstone();
  console.log("done");
})();
