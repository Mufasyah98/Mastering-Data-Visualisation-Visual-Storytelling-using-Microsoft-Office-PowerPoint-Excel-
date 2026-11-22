"""Generate the fictional Teratai Distribution datasets (H1 and Q3 2026)."""
import random, json, math
from datetime import date, timedelta

REGIONS = ["Central", "Northern", "Southern", "East Coast", "East Malaysia"]
CATS = ["Beverages", "Packaged Food", "Personal Care", "Household Care"]
LEAD = {"Central": 2, "Northern": 3, "Southern": 3, "East Coast": 4, "East Malaysia": 7}
CAT_MEAN = {"Beverages": 4500, "Packaged Food": 5200, "Personal Care": 3800, "Household Care": 6500}


def month_dates(y, m, n, rng):
    first = date(y, m, 1)
    nxt = date(y + (m == 12), m % 12 + 1, 1)
    days = (nxt - first).days
    return sorted(first + timedelta(days=rng.randrange(days)) for _ in range(n))


def make_orders(seed, months, start_id, region_w, ontime_p, late_max, cutoff):
    rng = random.Random(seed)
    rows = []
    oid = start_id
    for (y, m, n) in months:
        rw = region_w(m)
        for d in month_dates(y, m, n, rng):
            reg = rng.choices(REGIONS, weights=[rw[r] for r in REGIONS])[0]
            cat = rng.choices(CATS, weights=[30, 28, 22, 20])[0]
            val = round(CAT_MEAN[cat] * math.exp(rng.gauss(0, 0.45)) * 0.9, 2)
            val = max(650.0, min(val, 24800.0))
            prom = d + timedelta(days=LEAD[reg])
            if rng.random() < ontime_p(reg, m):
                act = prom - timedelta(days=rng.choice([0, 0, 0, 1]))
            else:
                act = prom + timedelta(days=rng.randint(1, late_max(reg, m)))
            status = "Delivered"
            if act > cutoff:
                act, status = None, "In Transit"
            rows.append(dict(order_id=f"ORD-26{oid:04d}", order_date=d, region=reg,
                             category=cat, value=val, promised=prom, actual=act, status=status))
            oid += 1
    return rows


# ---------- H1 (Jan-Jun 2026) ----------
H1_W = {"Central": 36, "Northern": 20, "Southern": 20, "East Coast": 10, "East Malaysia": 14}
H1_P = {"Central": .93, "Northern": .90, "Southern": .91, "East Coast": .89, "East Malaysia": .72}
h1 = make_orders(
    seed=2026,
    months=[(2026, 1, 45), (2026, 2, 44), (2026, 3, 50), (2026, 4, 52), (2026, 5, 53), (2026, 6, 56)],
    start_id=1,
    region_w=lambda m: H1_W,
    ontime_p=lambda r, m: H1_P[r],
    late_max=lambda r, m: 6 if r == "East Malaysia" else 3,
    cutoff=date(2026, 6, 30),
)

# ---------- Q3 (Jul-Sep 2026) : built-in pattern ----------
def q3_w(m):
    w = dict(H1_W)
    if m >= 8:  # Borneo promotion lifts East Malaysia volume
        w["East Malaysia"] = 19
    return w

def q3_p(r, m):
    if r == "East Malaysia":
        return {7: .76, 8: .52, 9: .42}[m]
    return {"Central": .93, "Northern": .91, "Southern": .91, "East Coast": .90}[r]

q3 = make_orders(
    seed=7309,
    months=[(2026, 7, 105), (2026, 8, 112), (2026, 9, 118)],
    start_id=301,
    region_w=q3_w,
    ontime_p=q3_p,
    late_max=lambda r, m: (8 if m >= 8 else 5) if r == "East Malaysia" else 3,
    cutoff=date(2026, 9, 30),
)


def ser(rows):
    out = []
    for r in rows:
        o = dict(r)
        for k in ("order_date", "promised", "actual"):
            o[k] = o[k].isoformat() if o[k] else None
        out.append(o)
    return out

json.dump({"h1": ser(h1), "q3": ser(q3)}, open("/home/claude/work/data.json", "w"), indent=0)
print(len(h1), len(q3))
