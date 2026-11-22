# Mastering Data Visualisation & Visual Storytelling using Microsoft Office (PowerPoint & Excel)

Hands-on exercise files for a two-day corporate workshop.

**Trainer:** Fakhrul Syahmi · Data Consultant & Certified Trainer · mufasyahcons@gmail.com
**Audience:** 20–24 corporate professionals with mixed Excel and PowerPoint experience
**Case:** Teratai Distribution Sdn Bhd, a fictional Malaysian FMCG distributor that tracks monthly order volume and delivery performance across regions (MYR)

## Files

| File | Session | Contents |
|---|---|---|
| `workshop-files/Day_1_PowerPoint_Exercises.pptx` | Day 1 | Ex 1 slide redesign · Ex 2 process infographic · Ex 3 chart selection. Each exercise has a brief, source content and working slides. |
| `workshop-files/Day_2_Excel_Exercises.xlsx` | Day 2 | Start Here · Raw Orders (306 rows with planted issues) · Cleaned Orders · Pivot Practice · Dashboard |
| `workshop-files/Capstone_Excel_Dataset.xlsx` | Day 2 capstone | Q3 2026 dataset (335 orders), management brief, analysis and mini-dashboard workspaces |
| `workshop-files/Capstone_PowerPoint_Template.pptx` | Day 2 capstone | Context / Insight / Recommendation slides + presentation and peer feedback checklist |
| `workshop-files/Trainer_Answer_Guide.docx` | Trainer only | Sample solutions, cleaning row counts, PivotTable mappings, KPI values, capstone worked example, scoring rubric |

> The Trainer Answer Guide contains the solutions. Do not share it with participants.

## Case definitions (consistent across all files)

- **Regions:** Central, Northern, Southern, East Coast, East Malaysia
- **Categories:** Beverages, Packaged Food, Personal Care, Household Care
- **Fields:** Order ID, Order Date, Region, Product Category, Order Value (MYR), Promised Delivery Date, Actual Delivery Date, Order Status
- **Lead time (calendar days):** Central 2 · Northern 3 · Southern 3 · East Coast 4 · East Malaysia 7
- **On-time delivery:** Actual Delivery Date on or before Promised Delivery Date
- **On-time rate:** on-time deliveries ÷ completed orders (Order Status = Delivered); in-transit orders excluded
- **Target:** 90% on time

## Key reference figures

| | Jan–Jun 2026 (Day 2) | Jul–Sep 2026 (Capstone) |
|---|---|---|
| Orders | 300 (after removing 6 duplicates) | 335 |
| Order value (MYR) | 1,525,604.27 | 1,700,542.37 |
| On-time rate | 85.5% (254/297) | 85.5% (271/317) |
| East Malaysia on-time rate | 70.7% (29/41) | 59.6% (34/57) |

## Build scripts

`build-scripts/` contains the generators used to produce the files (Python `openpyxl`/`pandas`, Node `pptxgenjs`/`docx`). Run order: `gen_data.py` → `build_xlsx.py` → `analyse.py` → `build_pptx.js` → `build_docx.js`. Data generation uses fixed random seeds, so the figures are reproducible. Paths inside the scripts point to the original working folder and need adjusting before reuse.

PivotTables, PivotCharts, slicers and timelines are not pre-built: the workbooks provide labelled workspaces and step-by-step instructions for participants to build them in Excel.

---
© 2026 Fakhrul Syahmi · mufasyahcons@gmail.com
