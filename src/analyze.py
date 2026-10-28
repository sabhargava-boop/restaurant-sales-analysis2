"""Restaurant sales-trend analysis: why did weekly sales fall?
Sales = Orders x Average check, so the change in sales splits into a traffic effect and a check-size effect.
Outputs percentages and indexed charts only (no dollar amounts), so results can be shared without exposing a business's data.

Usage:  python src/analyze.py                       # runs on the synthetic sample
        python src/analyze.py --input "path/Sales by day.csv" --end 2026-10-07
"""
import argparse
from pathlib import Path
import numpy as np, pandas as pd, matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

ap = argparse.ArgumentParser()
ap.add_argument("--input", default="sample_data/Sales by day.csv")
ap.add_argument("--out", default="docs")
ap.add_argument("--end", default=None, help="drop days after this date (e.g. an incomplete last day)")
ap.add_argument("--base", nargs=2, default=["2026-05-04", "2026-08-30"], help="base period first/last day (Mon..Sun)")
ap.add_argument("--recent", nargs=2, default=["2026-09-07", "2026-10-04"], help="recent period first/last day (Mon..Sun)")
a = ap.parse_args()
out = Path(a.out); out.mkdir(parents=True, exist_ok=True)

# ---- 1. load + clean ----
d = pd.read_csv(a.input)
d["date"] = pd.to_datetime(d["yyyyMMdd"].astype(str))
d = d.rename(columns={"Net sales": "net", "Total orders": "orders", "Total guests": "guests"}).drop(columns="yyyyMMdd")
if a.end: d = d[d.date <= a.end]
d["dow"] = d.date.dt.day_name()
d["week"] = d.date - pd.to_timedelta(d.date.dt.weekday, unit="D")

# ---- 2. full Mon-Sun weeks only ----
w = d.groupby("week").agg(days=("date", "count"), net=("net", "sum"), orders=("orders", "sum"), guests=("guests", "sum"))
w = w[w.days == 7].copy(); w["check"] = w.net / w.orders

b0, b1 = map(pd.Timestamp, a.base); r0, r1 = map(pd.Timestamp, a.recent)
base = w.loc[b0: b1 - pd.Timedelta(days=6)]; rec = w.loc[r0: r1 - pd.Timedelta(days=6)]
bn, rn = base.net.mean(), rec.net.mean(); bo, ro = base.orders.mean(), rec.orders.mean()
bc, rc = bn / bo, rn / ro

# ---- 3. decomposition (effects expressed as % of base weekly sales) ----
traffic = (ro - bo) * bc / bn * 100
check = ro * (rc - bc) / bn * 100
total = (rn / bn - 1) * 100
dow_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
db = d[(d.date >= b0) & (d.date <= b1)].groupby("dow").net.mean().reindex(dow_order)
dr = d[(d.date >= r0) & (d.date <= r1)].groupby("dow").net.mean().reindex(dow_order)
dow_pct = (dr / db - 1) * 100

summary = pd.Series({"weekly_sales_change_pct": total, "orders_change_pct": (ro / bo - 1) * 100,
                     "avg_check_change_pct": (rc / bc - 1) * 100, "effect_of_fewer_orders_pct_of_base": traffic,
                     "effect_of_check_size_pct_of_base": check, **{f"{k}_change_pct": v for k, v in dow_pct.items()}})
summary.round(1).to_csv(out / "summary.csv", header=["value"]); print(summary.round(1).to_string())

# ---- 4. charts (indexed: base-period average = 100) ----
plt.rcParams.update({"font.size": 11, "axes.spines.top": False, "axes.spines.right": False})
idx = w.net / bn * 100
fig, ax = plt.subplots(figsize=(9, 4.2))
ax.plot(idx.index, idx, marker="o", color="#9aa5b1", label="Weekly sales index")
ax.plot(idx.index, idx.rolling(4).mean(), lw=3, color="#c0392b", label="4-week average")
ax.axhline(100, color="black", lw=.8, ls="--"); ax.set_title("Weekly sales (base-period average = 100)"); ax.legend(frameon=False)
fig.autofmt_xdate(); fig.tight_layout(); fig.savefig(out / "weekly_sales_index.png", dpi=160); plt.close()

fig, ax = plt.subplots(figsize=(9, 4.2)); cols = ["#c0392b" if v < 0 else "#2e7d32" for v in dow_pct]
ax.bar([x[:3] for x in dow_order], dow_pct, color=cols); ax.axhline(0, color="black", lw=.8)
for i, v in enumerate(dow_pct): ax.annotate(f"{v:+.0f}%", (i, v), ha="center", va="bottom" if v >= 0 else "top")
ax.set_title("Change in average daily sales by weekday (recent vs. base)"); ax.set_ylabel("% change"); fig.tight_layout()
fig.savefig(out / "weekday_change.png", dpi=160); plt.close()

fig, ax = plt.subplots(figsize=(7, 4)); ax.bar(["Fewer orders", "Avg check"], [traffic, check], color=["#c0392b", "#2e7d32" if check >= 0 else "#c0392b"])
ax.axhline(0, color="black", lw=.8); ax.set_ylabel("% of base weekly sales"); ax.set_title(f"Where the {total:+.0f}% change came from")
for i, v in enumerate([traffic, check]): ax.annotate(f"{v:+.1f}", (i, v), ha="center", va="bottom" if v >= 0 else "top")
fig.tight_layout(); fig.savefig(out / "decomposition.png", dpi=160); plt.close()
