const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType, AlignmentType, HeadingLevel,
  Header, Footer, PageNumber, ImageRun, LevelFormat, BorderStyle, PageBreak, TableOfContents } = require("docx");

const S = JSON.parse(fs.readFileSync("/home/claude/work/stats.json"));
const I = JSON.parse(fs.readFileSync("/home/claude/work/issues.json"));
const REG = ["Central", "Northern", "Southern", "East Coast", "East Malaysia"];
const CATS = ["Beverages", "Packaged Food", "Personal Care", "Household Care"];
const H = S.h1, Q = S.q3, CL = S.cleaning;
const TEAL = "0F766E", ORANGE = "EA580C", LIGHT = "F0FDFA", INK = "1F2937";
const pct = (x, d = 1) => (x * 100).toFixed(d) + "%";
const n0 = (x) => Math.round(x).toLocaleString("en-US");
const n2 = (x) => Number(x).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const W = 9026; // A4 width minus 1" margins (DXA)

const P = (text, o = {}) => new Paragraph({ spacing: { after: 120 }, ...o.p, children: (Array.isArray(text) ? text : [text]).map((t) => typeof t === "string" ? new TextRun({ text: t, ...o.r }) : t) });
const B = (t) => new TextRun({ text: t, bold: true });
const bullet = (text) => new Paragraph({ numbering: { reference: "bul", level: 0 }, spacing: { after: 60 }, children: (Array.isArray(text) ? text : [text]).map((t) => typeof t === "string" ? new TextRun(t) : t) });
const num = (text, ref = "num") => new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 60 }, children: (Array.isArray(text) ? text : [text]).map((t) => typeof t === "string" ? new TextRun(t) : t) });
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(t)] });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(t)] });
const border = { style: BorderStyle.SINGLE, size: 4, color: "CBD5E1" };
const borders = { top: border, bottom: border, left: border, right: border };

function table(rows, widths, opt = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths,
    rows: rows.map((r, i) => new TableRow({ tableHeader: i === 0, children: r.map((v, j) => new TableCell({
      borders, width: { size: widths[j], type: WidthType.DXA }, margins: { top: 60, bottom: 60, left: 100, right: 100 },
      shading: { type: ShadingType.CLEAR, fill: i === 0 ? TEAL : (opt.hl && opt.hl.includes(i) ? "FFEDD5" : (i % 2 ? "FFFFFF" : "F8FAFC")), color: "auto" },
      children: String(v).split("\n").map((line) => new Paragraph({ alignment: (j > 0 && opt.right !== false && /^[\d.,%\-–MYR ×x()\/]+$/.test(String(v)) && i > 0) ? AlignmentType.RIGHT : AlignmentType.LEFT,
        children: [new TextRun({ text: line, bold: i === 0 || (opt.boldLast && i === rows.length - 1) || (opt.boldFirst && j === 0), color: i === 0 ? "FFFFFF" : INK, size: opt.size || 19 })] })),
    })) })) });
}
const gap = () => new Paragraph({ spacing: { after: 80 }, children: [] });
const img = (file, w = 600) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new ImageRun({ type: "png", data: fs.readFileSync(file), transformation: { width: w, height: Math.round(w * 0.5625) }, altText: { title: "Sample solution", description: "Sample slide", name: file } })] });
const caption = (t) => P(t, { r: { italics: true, size: 18, color: "64748B" }, p: { alignment: AlignmentType.CENTER } });
const note = (t) => new Paragraph({ spacing: { before: 80, after: 160 }, shading: { type: ShadingType.CLEAR, fill: LIGHT, color: "auto" }, border: { left: { style: BorderStyle.SINGLE, size: 4, color: LIGHT } },
  children: [new TextRun({ text: "Trainer note: ", bold: true, color: TEAL }), new TextRun(t)] });

// ---------- derived figures ----------
const em = Q.by_region["East Malaysia"];
const emM = ["Jul", "Aug", "Sep"].map((m) => Q.region_month["East Malaysia"][m]);
const sumK = (arr) => { const o = { orders: 0, value: 0, completed: 0, on_time: 0 }; arr.forEach((k) => { o.orders += k.orders; o.value += k.value; o.completed += k.completed; o.on_time += k.on_time; }); o.rate = o.on_time / o.completed; return o; };
const emAJ = sumK(["Apr", "May", "Jun"].map((m) => H.region_month["East Malaysia"][m]));
const h1EM = H.by_region["East Malaysia"];

const children = [];
// Title block
children.push(new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "Trainer Answer Guide", bold: true, size: 44, color: TEAL })] }));
children.push(P("Mastering Data Visualisation & Visual Storytelling using Microsoft Office (PowerPoint & Excel) · Two-day workshop", { r: { size: 24, color: INK } }));
children.push(P("Trainer: Fakhrul Syahmi  ·  Case: Teratai Distribution Sdn Bhd (fictional)  ·  Confidential: trainer use only", { r: { size: 20, color: "64748B" } }));

children.push(new Paragraph({ children: [new PageBreak()] }));

// 1
children.push(H1("1. Case constants and file map"));
children.push(P("Every workshop file uses the same company, regions, fields and KPI definitions. All figures in this guide were calculated from the delivered workbooks by repeating the participant cleaning steps exactly, and reconcile across files."));
children.push(table([["File", "Purpose", "Session"],
  ["Day_1_PowerPoint_Exercises.pptx", "Ex 1 slide redesign · Ex 2 process infographic · Ex 3 chart selection", "Day 1"],
  ["Day_2_Excel_Exercises.xlsx", "Clean data · PivotTables & PivotCharts · slicer & timeline · dashboard · chart to PowerPoint", "Day 2"],
  ["Capstone_Excel_Dataset.xlsx", "Q3 2026 dataset (Jul–Sep), management brief, mini-dashboard workspace", "Day 2 capstone"],
  ["Capstone_PowerPoint_Template.pptx", "Context / Insight / Recommendation slides + feedback checklist", "Day 2 capstone"],
  ["Trainer_Answer_Guide.docx", "This guide", "Trainer"]], [3100, 4626, 1300]));
children.push(gap());
children.push(table([["Item", "Value used in every file"],
  ["Company", "Teratai Distribution Sdn Bhd — FMCG distributor, DC in Shah Alam (fictional)"],
  ["Regions (5)", REG.join(", ")],
  ["Product categories (4)", CATS.join(", ")],
  ["Standard lead time", "Central 2 · Northern 3 · Southern 3 · East Coast 4 · East Malaysia 7 calendar days"],
  ["On-time delivery", "Actual Delivery Date on or before Promised Delivery Date"],
  ["On-time rate", "On-time deliveries ÷ completed orders (Order Status = Delivered); In Transit excluded"],
  ["On-time target", "90%"],
  ["Order IDs", "H1: ORD-260001 to ORD-260300 · Q3: ORD-260301 to ORD-260635 (no overlap)"],
  ["Currency", "MYR, two decimals in workbooks"]], [2600, 6426]));

// 2 Day 1
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(H1("2. Day 1 — PowerPoint sample solutions"));
children.push(H2("Exercise 1 · Slide redesign"));
children.push(P([B("Main finding: "), `East Malaysia has the weakest on-time delivery rate, ${pct(h1EM.rate)} (${h1EM.on_time} of ${h1EM.completed} completed orders), against ${pct(S.h1_others.rate)} for all other regions combined and a 90% target. East Malaysia has ${pct(S.h1_em_order_share, 0)} of orders but ${pct(S.h1_em_late_share, 0)} of late deliveries (${h1EM.late} of ${H.total.late}).`]));
children.push(P([B("Suggested headline: "), `“East Malaysia delivers only ${pct(h1EM.rate, 0)} of orders on time — well below the 90% target”.`]));
children.push(P(B("What good redesigns do:")));
["Keep one chart: On-Time % only, sorted, East Malaysia in the accent colour and the other regions grey. Delete the Orders and Late Deliveries series (different units on one axis).",
  "Remove the yellow plot area, gridlines, legend and the 'NEW!' star; turn on data labels with one decimal.",
  "Delete the stock-take reminder and the new-hire note (not about this message). Move 'Central = biggest region' and 'Packaged Food' to the appendix or speaker notes.",
  "Turn the remaining facts into 2–3 supporting callouts: EM rate, all-other-regions rate, EM share of late deliveries.",
  "Use one font, a left-aligned headline and a source line in small grey text."].forEach((t) => children.push(bullet(t)));
children.push(img("/home/claude/work/qa/sol-1.png"));
children.push(caption("Sample redesign of the Exercise 1 slide"));
children.push(note("Common mistake: participants keep all three series and only change the colours. Ask them, “What do you want the Director to notice?” and then, “Which series proves it?”"));

children.push(H2("Exercise 2 · Process infographic"));
["Best SmartArt layouts: Basic Chevron Process, Continuous Block Process or Step Up Process. Convert to SmartArt works directly on the bulleted working slide.",
  "Step names stay short (2–4 words); owner and time go in a smaller grey label under each step.",
  "One icon per step: cart (order received), clipboard (credit & stock check), box (pick & pack), route (dispatch), truck (delivery), invoice (proof of delivery).",
  "Optional emphasis: highlight Step 5 (delivery) in the accent colour — it links to the on-time KPI used on Day 2.",
  "Show the lead-time rule once, at the bottom."].forEach((t) => children.push(bullet(t)));
children.push(img("/home/claude/work/qa/sol-2.png"));
children.push(caption("Sample process infographic built with editable shapes and icons"));

children.push(H2("Exercise 3 · Chart selection"));
const t3 = [["Question", "Recommended chart", "Avoid", "Sample takeaway"],
  ["A · Which region contributes the most order value?", "Sorted horizontal bar (ranking)", "Pie with 5 slices; unsorted columns", `Central brings in ${pct(H.by_region.Central.value / H.total.value, 0)} of H1 order value (MYR ${n0(H.by_region.Central.value)}), about ${(H.by_region.Central.value / H.by_region.Southern.value).toFixed(1)}× Southern (MYR ${n0(H.by_region.Southern.value)}).`],
  ["B · How did monthly order volume change?", "Line chart (or columns) by month", "Bars sorted by size (breaks time order)", `Monthly orders rose ${pct(H.by_month.Jun.orders / H.by_month.Jan.orders - 1, 0)}, from ${H.by_month.Jan.orders} in January to ${H.by_month.Jun.orders} in June; order value peaked in May (MYR ${n0(H.by_month.May.value)}).`],
  ["C · Which region has the weakest on-time delivery?", "Sorted bar of on-time rate with 90% target line, or 100% stacked bar (on time vs late)", "Clustered bars of counts (regions differ in size)", `Only Southern meets the 90% target; East Malaysia is lowest at ${pct(h1EM.rate)} (${h1EM.on_time} of ${h1EM.completed}).`]];
children.push(table(t3, [2100, 2100, 1700, 3126]));
children.push(gap());
children.push(img("/home/claude/work/qa/sol-3.png"));
children.push(caption("Sample answers for Exercise 3"));
children.push(note(`Table C shows ${H.total.in_transit} in-transit orders. Check that participants exclude them from the rate: ${H.total.on_time} ÷ ${H.total.completed} = ${pct(H.total.rate)}, not ${H.total.on_time} ÷ ${H.total.orders} = ${pct(H.total.on_time / H.total.orders)}.`));

// 3 Day 2 cleaning
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(H1("3. Day 2 — Data cleaning: expected steps and row counts"));
children.push(table([["Issue planted in 'Raw Orders'", "Rows affected", "Fix"],
  ["Duplicate Order IDs (exact duplicate rows)", String(CL.duplicates), "Data ▸ Remove Duplicates on Order ID"],
  ["Trailing spaces in Region", String(CL.trail_region), "TRIM, then Paste Values"],
  ["Trailing spaces in Product Category", String(CL.trail_cat), "TRIM, then Paste Values"],
  ["Inconsistent Region labels", String(CL.region_variants_after_trim), "Find & Replace (match entire cell)"],
  ["Order Date stored as text (yyyy-mm-dd)", String(CL.text_order_dates), "DATEVALUE or Text to Columns (YMD)"],
  ["Actual Delivery Date stored as text", String(CL.text_actual_dates), "DATEVALUE or Text to Columns (YMD)"]], [4000, 1400, 3626]));
children.push(gap());
children.push(P(B("Region label variants and their standard names")));
const canon = { central: "Central", CENTRAL: "Central", "Northern Region": "Northern", southern: "Southern", "E. Coast": "East Coast", "East Msia": "East Malaysia" };
children.push(table([["Variant in raw data", "Rows", "Replace with"], ...Object.entries(I.variants).map(([k, v]) => [`“${k}”`, String(v), canon[k]])], [3500, 1200, 4326]));
children.push(gap());
children.push(P(B("Row-count checkpoints")));
children.push(table([["Step", "Check", "Expected result"],
  ["1 Paste values", "Data rows", String(CL.raw_rows)],
  ["2 Remove duplicates", "Removed / remaining", `${CL.duplicates} removed · ${CL.after_dedup} remaining`],
  ["", "Duplicate IDs", CL.dup_ids.join(", ")],
  ["3–4 Trim & standardise", "Distinct Region values", `${CL.raw_distinct_region_labels} before → 5 after`],
  ["", "Distinct Product Category values", "8 before (4 with trailing spaces) → 4 after"],
  ["5 Convert dates", "Text dates remaining", `${CL.text_order_dates + CL.text_actual_dates} before → 0 after`],
  ["7 Helper columns", "Completed orders (Completed = 1)", `${H.total.completed} (${H.total.in_transit} In Transit)`],
  ["", "Sum of On-Time Flag", String(H.total.on_time)],
  ["", "Total Order Value (MYR)", n2(H.total.value)]], [2400, 3000, 3626]));
children.push(note("Order of steps matters only for Remove Duplicates if participants choose 'all columns': the duplicates are exact copies, so either approach removes 6 rows. If someone gets 299 or 301 rows, check whether they included the header or removed duplicates on the wrong column."));

// 4 Pivot
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(H1("4. PivotTable field mappings and expected results (Jan–Jun 2026)"));
children.push(table([["#", "Rows", "Columns", "Values", "Filter"],
  ["Q1", "Region", "—", "Sum of Order Value (MYR)", "—"],
  ["Q2", "Order Date (Months)", "—", "Count of Order ID; Sum of Order Value", "—"],
  ["Q3", "Region", "Product Category", "Sum of Order Value (% of Row Total)", "—"],
  ["Q4", "Region", "—", "Average of On-Time Flag", "Order Status = Delivered"],
  ["Q5", "Order Date (Months)", "—", "Average of On-Time Flag; Count of Order ID", "Order Status = Delivered"],
  ["Q6", "Product Category", "—", "Average of Order Value (MYR)", "—"],
  ["Q7", "Region", "—", "Average / Max of Days Late", "Delivered; On-Time Flag = 0"]], [600, 1900, 1600, 3026, 1900], { right: false }));
children.push(gap());
const byVal = [...REG].sort((a, b) => H.by_region[b].value - H.by_region[a].value);
children.push(P(B("Q1 · Order value by region (sorted) and Q4 · on-time rate by region")));
children.push(table([["Region", "Orders", "Order Value (MYR)", "Share", "Completed", "On time", "On-time rate"],
  ...byVal.map((r) => { const k = H.by_region[r]; return [r, String(k.orders), n2(k.value), pct(k.value / H.total.value), String(k.completed), String(k.on_time), pct(k.rate)]; }),
  ["Grand Total", String(H.total.orders), n2(H.total.value), "100.0%", String(H.total.completed), String(H.total.on_time), pct(H.total.rate)]], [1600, 900, 1700, 1000, 1200, 1100, 1526], { boldLast: true }));
children.push(gap());
children.push(P(B("Q2 · Monthly orders and value; Q5 · on-time rate by month")));
children.push(table([["Month", "Orders", "Order Value (MYR)", "Completed", "On time", "On-time rate"],
  ...H.months.map((m) => { const k = H.by_month[m]; return [m + " 2026", String(k.orders), n2(k.value), String(k.completed), String(k.on_time), pct(k.rate)]; }),
  ["Grand Total", String(H.total.orders), n2(H.total.value), String(H.total.completed), String(H.total.on_time), pct(H.total.rate)]], [1600, 1100, 2000, 1400, 1300, 1626], { boldLast: true }));
children.push(gap());
children.push(P(B("Q3 · Order value by region and category (% of row total)")));
children.push(table([["Region", ...CATS, "Total (MYR)"], ...REG.map((r) => { const v = H.region_cat_value[r]; const t = CATS.reduce((a, c) => a + v[c], 0); return [r, ...CATS.map((c) => pct(v[c] / t)), n0(t)]; })], [1700, 1450, 1500, 1450, 1500, 1426]));
children.push(gap());
children.push(P(B("Q6 · Average order value by category and Q7 · lateness by region")));
const byAov = [...CATS].sort((a, b) => H.by_cat[b].aov - H.by_cat[a].aov);
children.push(table([["Category", "Orders", "Order Value (MYR)", "Average (MYR)"], ...byAov.map((c) => [c, String(H.by_cat[c].orders), n2(H.by_cat[c].value), n2(H.by_cat[c].aov)])], [2600, 1400, 2600, 2426]));
children.push(gap());
children.push(table([["Region", "Late orders", "Average days late", "Max days late"], ...REG.map((r) => [r, String(H.by_region[r].late), H.by_region[r].avg_days_late.toFixed(2), String(H.by_region[r].max_days_late)]),
  ["All regions", String(H.total.late), H.total.avg_days_late.toFixed(2), String(H.total.max_days_late)]], [2600, 1800, 2400, 2226], { boldLast: true }));
children.push(note("Average of On-Time Flag equals the on-time rate only when In Transit orders are filtered out (or their flag is blank). If a participant reports 84.7% instead of 85.5%, they divided by all 300 orders."));

// 5 Dashboard
children.push(H1("5. Dashboard KPI values and slicer tests"));
children.push(table([["KPI", "No filter (Jan–Jun)", "Slicer: East Malaysia", "East Malaysia + timeline Apr–Jun"],
  ["Total Orders", String(H.total.orders), String(h1EM.orders), String(emAJ.orders)],
  ["Total Order Value (MYR)", n2(H.total.value), n2(h1EM.value), n2(emAJ.value)],
  ["Average Order Value (MYR)", n2(H.total.aov), n2(h1EM.aov), n2(emAJ.value / emAJ.orders)],
  ["On-Time Delivery Rate", `${pct(H.total.rate)} (${H.total.on_time}/${H.total.completed})`, `${pct(h1EM.rate)} (${h1EM.on_time}/${h1EM.completed})`, `${pct(emAJ.rate)} (${emAJ.on_time}/${emAJ.completed})`]], [2600, 2200, 2000, 2226]));
children.push(gap());
children.push(P(B("Usability checks when you walk the room")));
["KPI cards use GETPIVOTDATA (or a direct link) to a PivotTable, so they respond to the slicer and timeline.",
  "Slicer and timeline are connected to every PivotTable (Report Connections).",
  "PivotChart field buttons hidden; chart titles state the message, not just the metric.",
  "Headline in row 4 is a sentence with a number, e.g. “East Malaysia's on-time rate (70.7%) is the main gap against the 90% target.”"].forEach((t) => children.push(bullet(t)));

children.push(H2("Exercise 5 · Linking vs embedding (expected answers)"));
children.push(table([["Question", "Linked", "Embedded"],
  ["Updates when Excel data changes?", "Yes, after Refresh Data (or on open, if the link path is valid)", "No — keeps a copy of the data at the time of pasting"],
  ["Email only the .pptx?", "Recipient sees the last image but cannot refresh; link breaks", "Works anywhere; data travels inside the file"],
  ["File size", "Small", "Larger (a copy of the workbook is stored)"],
  ["Edit chart data in PowerPoint?", "Opens the source Excel file", "Edit Data opens the embedded copy"],
  ["Typical use", "Monthly reports rebuilt from the same workbook", "Final decks sent to management or clients"]], [2600, 3200, 3226]));

// 6 Capstone
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(H1("6. Capstone — worked example (Q3 2026)"));
children.push(H2("Built-in pattern"));
children.push(P(`The Q3 dataset (${S.q3_rows} orders, ORD-260301 to ORD-260635) contains a deliberate pattern: East Malaysia on-time delivery deteriorates sharply after July, while every other region stays at about 91%. Because East Malaysia is a small region, the company-wide rate (${pct(Q.total.rate)}) looks the same as H1 (${pct(H.total.rate)}) — the problem is hidden in the total and only appears when groups slice by region and month.`));
children.push(table([["KPI (mini dashboard)", "Q3 value"],
  ["Total Orders", String(Q.total.orders)], ["Total Order Value (MYR)", n2(Q.total.value)], ["On-Time Delivery Rate", `${pct(Q.total.rate)} (${Q.total.on_time} of ${Q.total.completed} completed; ${Q.total.in_transit} in transit)`]], [4000, 5026]));
children.push(gap());
children.push(table([["Region", "Orders", "Order Value (MYR)", "Completed", "On time", "On-time rate", "Avg days late"],
  ...REG.map((r) => { const k = Q.by_region[r]; return [r, String(k.orders), n2(k.value), String(k.completed), String(k.on_time), pct(k.rate), k.avg_days_late.toFixed(2)]; }),
  ["Total", String(Q.total.orders), n2(Q.total.value), String(Q.total.completed), String(Q.total.on_time), pct(Q.total.rate), Q.total.avg_days_late.toFixed(2)]], [1500, 900, 1800, 1150, 1000, 1300, 1376], { boldLast: true, hl: [5] }));
children.push(gap());
children.push(table([["Month", "EM orders", "EM completed", "EM on time", "EM rate", "Other regions rate", "All regions rate"],
  ...["Jul", "Aug", "Sep"].map((m, i) => [m + " 2026", String(emM[i].orders), String(emM[i].completed), String(emM[i].on_time), pct(emM[i].rate), pct(S.q3_others_by_month[m].rate), pct(Q.by_month[m].rate)])], [1300, 1200, 1300, 1200, 1100, 1500, 1426]));
children.push(gap());
children.push(table([["East Malaysia by category", "Completed", "On time", "On-time rate"], ...CATS.map((c) => { const k = Q.cat_region["East Malaysia"][c]; return [c, String(k.completed), String(k.on_time), pct(k.rate)]; })], [3500, 1800, 1800, 1926]));
children.push(gap());
children.push(H2("Sample presentation"));
children.push(P([B("Context: "), "Complaints about late deliveries rose in Q3 and the Board meets in two weeks. Management must decide where to focus delivery improvement first."]));
children.push(P([B("Insight headline: "), `“East Malaysia on-time delivery fell from ${pct(emM[0].rate, 0)} in July to ${pct(emM[2].rate, 0)} in September, while other regions held at ~91%.” Evidence: East Malaysia ${pct(em.rate)} for the quarter (${em.on_time}/${em.completed}); ${pct(S.q3_em_late_share, 0)} of all late deliveries (${em.late} of ${Q.total.late}) from ${pct(S.q3_em_order_share, 0)} of orders; East Malaysia volume rose from ${emM[0].orders} to ${emM[2].orders} orders per month.`]));
children.push(img("/home/claude/work/qa/sol-4.png"));
children.push(caption("Sample Insight slide for the capstone"));
children.push(P([B("Recommendation: "), "Hold an urgent performance review with the East Malaysia 3PL / sea-freight partner within two weeks, agree a recovery plan, and track East Malaysia on-time rate weekly with a target of 85% by December 2026 (owner: Transport Planning Manager). Do not change lead times for other regions — they are performing on target."]));
children.push(H2("What the data supports — and what needs further investigation"));
children.push(P(B("Supported by the data")));
[`East Malaysia is the region where delivery performance needs attention: ${pct(em.rate)} vs ${pct(S.q3_others.rate)} for the other four regions combined.`,
  `The decline is concentrated in August and September (${pct(emM[1].rate)} and ${pct(emM[2].rate)}), after ${pct(emM[0].rate)} in July.`,
  "The problem is spread across all four categories in East Malaysia — it is a regional, not a product, issue.",
  "Other regions are stable month to month, so the company-wide average hides the problem."].forEach((t) => children.push(bullet(t)));
children.push(P(B("Requires further investigation (not in the dataset)")));
["Root cause: the data has no carrier, shipping-route, port or weather field. Groups cannot prove the 3PL is responsible — only that late deliveries are concentrated in East Malaysia.",
  `September is incomplete: ${emM[2].orders - emM[2].completed} East Malaysia orders were still In Transit at 30 Sep, so its rate may change.`,
  `Small samples: monthly East Malaysia rates rest on about 20 completed orders each; category splits inside East Malaysia (11–17 orders) are too small to rank categories.`,
  `Southern dipped to ${pct(Q.by_region.Southern.rate)} (${Q.by_region.Southern.on_time}/${Q.by_region.Southern.completed}) in Q3 from ${pct(H.by_region.Southern.rate)} in H1 — worth monitoring, but the evidence is weaker than for East Malaysia.`,
  "Whether the 7-day East Malaysia lead time is realistic: groups may suggest reviewing the promise as well as the execution."].forEach((t) => children.push(bullet(t)));
children.push(note("Reward groups that name the limitation (no carrier field) instead of claiming the 3PL is the cause. A recommendation to 'investigate East Malaysia delivery' is acceptable only if it has an owner, a deadline and a KPI."));

// 7 Reconciliation
children.push(H1("7. Cross-file reconciliation"));
children.push(table([["Figure", "Where it appears", "Value", "Status"],
  ["H1 orders after cleaning", "Day 2 checkpoints · Day 1 tables A–C · Pivot grand totals", String(H.total.orders), "Reconciled"],
  ["H1 order value (MYR)", "Day 1 table A & clutter slide · Pivot Q1/Q2 · Dashboard KPI", n2(H.total.value), "Reconciled"],
  ["H1 on-time rate", "Day 1 table C & clutter slide · Pivot Q4/Q5 · Dashboard KPI", `${pct(H.total.rate)} (${H.total.on_time}/${H.total.completed})`, "Reconciled"],
  ["H1 East Malaysia on-time rate", "Day 1 Ex 1 & Ex 3 · Pivot Q4", `${pct(h1EM.rate)} (${h1EM.on_time}/${h1EM.completed})`, "Reconciled"],
  ["Q3 orders / value", "Capstone brief & data · mini dashboard", `${Q.total.orders} / ${n2(Q.total.value)}`, "Reconciled"],
  ["Order ID ranges", "Day 2 and Capstone workbooks", "ORD-260001–260300 · ORD-260301–260635", "No overlap"]], [2300, 3400, 2100, 1226], { right: false }));

// 8 Rubric
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(H1("8. Capstone scoring rubric"));
children.push(P("Score each criterion 1–4. Each criterion carries 25% (maximum 16 points; multiply by 6.25 for a percentage)."));
children.push(table([["Criterion", "4 · Excellent", "3 · Good", "2 · Developing", "1 · Beginning"],
  ["Data accuracy", "All KPIs correct (335 orders; MYR 1,700,542.37; 85.5%); in-transit excluded; counts shown next to rates", "KPIs correct; minor labelling or rounding issue", "One KPI wrong (e.g. rate over all orders = 80.9%)", "Several figures wrong or unverifiable"],
  ["Dashboard usability", "3 KPIs, 2 charts and a region slicer all connected; clean layout; titles state the message", "All elements present; minor layout or connection issue", "Missing one element or slicer not connected to all PivotTables", "Several elements missing; hard to read"],
  ["Storytelling", "Context–Insight–Recommendation flow; takeaway headline with a number; highlight colour guides the eye", "Clear flow; headline is a topic label rather than a finding", "Slides are descriptive; message must be inferred", "No clear message"],
  ["Quality of recommendation", "One specific action (owner, timing, KPI target) that follows from the East Malaysia evidence; data limitations acknowledged", "Action follows from evidence but lacks owner or measure", "Generic action ('improve delivery') or weakly linked to findings", "No recommendation, or contradicts the evidence"]], [1600, 2050, 1800, 1800, 1776], { right: false, boldFirst: true, size: 17 }));
children.push(gap());
children.push(P(`Check value for the 'over all orders' mistake: ${Q.total.on_time} ÷ ${Q.total.orders} = ${pct(Q.total.on_time / Q.total.orders)}.`, { r: { italics: true, color: "64748B", size: 18 } }));

const doc = new Document({
  creator: "Fakhrul Syahmi", title: "Trainer Answer Guide",
  styles: { default: { document: { run: { font: "Arial", size: 21 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: "Arial", size: 30, bold: true, color: TEAL }, paragraph: { spacing: { before: 300, after: 140 }, keepNext: true, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: "Arial", size: 24, bold: true, color: ORANGE }, paragraph: { spacing: { before: 220, after: 100 }, keepNext: true, outlineLevel: 1 } }] },
  numbering: { config: [
    { reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] },
    { reference: "num", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1300, bottom: 1200, left: 1440, right: 1440 } } },
    headers: { default: new Header({ children: [new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: ORANGE, space: 4 } },
      children: [new TextRun({ text: "Fakhrul Syahmi | Data Consultant & Certified Trainer", bold: true, color: TEAL, size: 18 })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
      new TextRun({ text: "© 2026 Fakhrul Syahmi · mufasyahcons@gmail.com · Page ", size: 16, color: "64748B" }), new TextRun({ children: [PageNumber.CURRENT], size: 16, color: "64748B" })] })] }) },
    children }],
});
Packer.toBuffer(doc).then((b) => { fs.writeFileSync("/home/claude/work/out/Trainer_Answer_Guide.docx", b); console.log("ok"); });
