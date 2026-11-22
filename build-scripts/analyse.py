"""Re-read the delivered workbooks, clean exactly as participants will, compute every figure."""
import json, datetime as dt
import pandas as pd
from openpyxl import load_workbook

REG = ["Central", "Northern", "Southern", "East Coast", "East Malaysia"]
CATS = ["Beverages", "Packaged Food", "Personal Care", "Household Care"]
MAP = {"central": "Central", "CENTRAL": "Central", "Northern Region": "Northern", "southern": "Southern",
       "E. Coast": "East Coast", "East Msia": "East Malaysia"}
out = {}

# ---------- Day 2 raw ----------
wb = load_workbook("out/Day_2_Excel_Exercises.xlsx")
ws = wb["Raw Orders"]
vals = list(ws.iter_rows(min_row=1, max_col=8, values_only=True))
raw = pd.DataFrame(vals[1:], columns=vals[0])
n_raw = len(raw)
# issue audit on the raw sheet itself
text_od = sum(isinstance(v, str) for v in raw["Order Date"])
text_ad = sum(isinstance(v, str) for v in raw["Actual Delivery Date"])
trail_reg = sum(isinstance(v, str) and v != v.rstrip() for v in raw["Region"])
trail_cat = sum(isinstance(v, str) and v != v.rstrip() for v in raw["Product Category"])
dups = raw.duplicated("Order ID").sum()
c = raw.drop_duplicates("Order ID").copy()
c["Region"] = c["Region"].str.strip()
c["Product Category"] = c["Product Category"].str.strip()
bad_region = (~c["Region"].isin(REG)).sum()
c["Region"] = c["Region"].replace(MAP)
for col in ["Order Date", "Promised Delivery Date", "Actual Delivery Date"]:
    c[col] = pd.to_datetime(c[col].apply(lambda v: v if not isinstance(v, str) else dt.datetime.fromisoformat(v)))
assert set(c["Region"]) == set(REG) and set(c["Product Category"]) == set(CATS)
raw_distinct_regions = sorted(set(raw["Region"]))
out["cleaning"] = dict(raw_rows=n_raw, duplicates=int(dups), after_dedup=len(c), trail_region=trail_reg,
                       trail_cat=trail_cat, region_variants_after_trim=int(bad_region), text_order_dates=text_od,
                       text_actual_dates=text_ad, raw_distinct_region_labels=len(raw_distinct_regions),
                       raw_region_labels=raw_distinct_regions,
                       dup_ids=sorted(raw.loc[raw.duplicated("Order ID"), "Order ID"].tolist()))


def kpis(df):
    comp = df[df["Order Status"] == "Delivered"]
    ont = (comp["Actual Delivery Date"] <= comp["Promised Delivery Date"])
    late = comp[~ont]
    return dict(orders=len(df), value=round(df["Order Value (MYR)"].sum(), 2),
                aov=round(df["Order Value (MYR)"].mean(), 2), completed=len(comp), in_transit=len(df) - len(comp),
                on_time=int(ont.sum()), late=int((~ont).sum()), rate=ont.mean(),
                avg_days_late=round(((late["Actual Delivery Date"] - late["Promised Delivery Date"]).dt.days).mean(), 2) if len(late) else 0,
                max_days_late=int(((late["Actual Delivery Date"] - late["Promised Delivery Date"]).dt.days).max()) if len(late) else 0)


def breakdown(df):
    res = {"total": kpis(df)}
    res["by_region"] = {r: kpis(df[df["Region"] == r]) for r in REG}
    df = df.assign(M=df["Order Date"].dt.strftime("%b"))
    months = list(dict.fromkeys(df.sort_values("Order Date")["M"]))
    res["months"] = months
    res["by_month"] = {m: kpis(df[df["M"] == m]) for m in months}
    res["by_cat"] = {k: kpis(df[df["Product Category"] == k]) for k in CATS}
    res["region_cat_value"] = {r: {k: round(df[(df.Region == r) & (df["Product Category"] == k)]["Order Value (MYR)"].sum(), 2) for k in CATS} for r in REG}
    res["region_month"] = {r: {m: kpis(df[(df.Region == r) & (df.M == m)]) for m in months} for r in REG}
    res["cat_region"] = {r: {k: kpis(df[(df.Region == r) & (df["Product Category"] == k)]) for k in CATS} for r in REG}
    return res

out["h1"] = breakdown(c)

# ---------- Capstone ----------
q = pd.read_excel("out/Capstone_Excel_Dataset.xlsx", sheet_name="Orders Q3")
for col in ["Order Date", "Promised Delivery Date", "Actual Delivery Date"]:
    q[col] = pd.to_datetime(q[col])
out["q3_rows"] = len(q)
out["q3"] = breakdown(q)
# others combined
oth = q[q.Region != "East Malaysia"]
out["q3_others"] = kpis(oth)
out["q3_others_by_month"] = {m: kpis(oth[oth["Order Date"].dt.strftime("%b") == m]) for m in ["Jul", "Aug", "Sep"]}
# late share
tk = out["q3"]["total"]
em = out["q3"]["by_region"]["East Malaysia"]
out["q3_em_late_share"] = em["late"] / tk["late"]
out["q3_em_order_share"] = em["orders"] / tk["orders"]
json.dump(out, open("stats.json", "w"), indent=1, default=float)

# ---- print summary ----
def p(name, k):
    print(f"{name:15s} orders {k['orders']:4d} value {k['value']:12,.2f} aov {k['aov']:9,.2f} comp {k['completed']:4d} ontime {k['on_time']:4d} rate {k['rate']:.1%} late {k['late']} avgLate {k['avg_days_late']}")
print(out["cleaning"])
for key in ["h1", "q3"]:
    print("=====", key)
    s = out[key]
    p("TOTAL", s["total"])
    for r in REG: p(r, s["by_region"][r])
    for m in s["months"]: p(m, s["by_month"][m])
    for k in CATS: p(k, s["by_cat"][k])
print("others", out["q3_others"]["rate"], {m: round(v["rate"], 3) for m, v in out["q3_others_by_month"].items()})
for m in ["Jul", "Aug", "Sep"]:
    print("EM", m, out["q3"]["region_month"]["East Malaysia"][m]["on_time"], "/", out["q3"]["region_month"]["East Malaysia"][m]["completed"], out["q3"]["region_month"]["East Malaysia"][m]["orders"])
for k in CATS:
    e = out["q3"]["cat_region"]["East Malaysia"][k]; print("EM", k, e["on_time"], e["completed"], round(e["rate"], 3))
print("EM late share", out["q3_em_late_share"], "order share", out["q3_em_order_share"])
oh = c[c.Region != "East Malaysia"]
out["h1_others"] = kpis(oh)
out["h1_em_late_share"] = out["h1"]["by_region"]["East Malaysia"]["late"] / out["h1"]["total"]["late"]
out["h1_em_order_share"] = out["h1"]["by_region"]["East Malaysia"]["orders"] / out["h1"]["total"]["orders"]
json.dump(out, open("stats.json", "w"), indent=1, default=float)
p("H1 others", out["h1_others"]); print(out["h1_em_late_share"], out["h1_em_order_share"])
