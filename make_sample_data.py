"""Generate demo sales files (one per branch, with some messy rows) for testing."""

import random
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

random.seed(7)
PRODUCTS = {"Carpet Tiles": 120, "Vinyl Flooring": 85, "Rubber Mat": 40, "Artificial Grass": 65, "SPC Flooring": 150}
BRANCHES = ["Cairo", "Giza", "Alexandria", "Mansoura"]

out = Path("sample_data")
out.mkdir(exist_ok=True)
start = date(2026, 10, 1)

for branch in BRANCHES:
    rows = []
    for _ in range(random.randint(60, 90)):
        product = random.choice(list(PRODUCTS))
        rows.append({
            "date": start + timedelta(days=random.randint(0, 29)),
            "branch": f" {branch} ",            # messy spaces on purpose
            "product": product,
            "quantity": random.randint(1, 15),
            "unit price": PRODUCTS[product],
        })
    rows.append(dict(rows[0]))                   # duplicate row
    rows.append({"date": "n/a", "branch": branch, "product": "Rubber Mat", "quantity": "x", "unit price": 40})
    df = pd.DataFrame(rows)
    if branch == "Giza":
        df.to_csv(out / f"{branch.lower()}_sales.csv", index=False)
    else:
        df.to_excel(out / f"{branch.lower()}_sales.xlsx", index=False)

print(f"Sample files created in {out}/")
