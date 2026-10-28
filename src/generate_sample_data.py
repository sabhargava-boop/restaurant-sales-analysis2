"""Generate SYNTHETIC restaurant 'Sales by day' data (invented numbers, same column format as a POS export).
A weeknight slowdown is deliberately built into September so the analysis has something to find.
Nothing here comes from a real business."""
import numpy as np, pandas as pd
from pathlib import Path

rng = np.random.default_rng(7)
dates = pd.date_range("2026-05-01", "2026-10-07")
base_orders = {0: 21, 1: 19, 2: 20, 3: 22, 4: 33, 5: 37, 6: 30}     # Mon=0 ... Sun=6
slowdown   = {0: .91, 1: .80, 2: .87, 3: .83, 4: 1.0, 5: 1.02, 6: .90}  # applied from Sep 1
rows = []
for d in dates:
    f = slowdown[d.weekday()] if d >= pd.Timestamp("2026-09-01") else 1.0
    orders = max(5, int(rng.normal(base_orders[d.weekday()] * f, 3)))
    check = float(np.clip(rng.normal(41, 3.5), 28, 55))
    guests = int(orders * rng.normal(1.6, 0.15))
    rows.append((d.strftime("%Y%m%d"), round(orders * check, 2), orders, guests))
out = Path("sample_data"); out.mkdir(exist_ok=True)
pd.DataFrame(rows, columns=["yyyyMMdd", "Net sales", "Total orders", "Total guests"]).to_csv(out / "Sales by day.csv", index=False)
print("wrote", out / "Sales by day.csv", len(rows), "days")
