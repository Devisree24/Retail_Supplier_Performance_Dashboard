"""
Retail Supplier Performance & Inventory Optimization - Data Generator
P&G supplies to Walmart, Sam's Club, Costco. Overstock = cash locked + storage cost;
Understock = lost sales + retailer penalties. Pricing & promotions affect sell-through.

Generates 80,000-120,000 rows across Sales, Inventory, and Pricing & Promo tables.
Run: python scripts/generate_data.py
Output: data/Sales.csv, data/Inventory.csv, data/Pricing_Promo.csv
"""

import csv
import random
from datetime import datetime, timedelta
from pathlib import Path

# --- Configuration ---
RANDOM_SEED = 42
random.seed(RANDOM_SEED)

# Target: 80,000 - 120,000 sales rows
NUM_SALES_ROWS = 105_000
NUM_SKUS = 220
DAYS_OF_DATA = 365

WALMART_STORES = 48
SAMS_STORES = 28
COSTCO_STORES = 18

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = PROJECT_ROOT / "data"
DATA_DIR.mkdir(exist_ok=True)

# Product categories per your spec: household, personal care, packaged foods, beverages, etc.
CATEGORIES = [
    "household",
    "personal care",
    "packaged foods",
    "beverages",
    "baby care",
    "cleaning",
]

REGIONS = ["Midwest", "South", "West", "Northeast"]

def generate_skus():
    """Generate SKU list with product category."""
    skus = []
    cats_short = {
        "household": "HH",
        "personal care": "PC",
        "packaged foods": "PF",
        "beverages": "BV",
        "baby care": "BC",
        "cleaning": "CL",
    }
    for i in range(1, NUM_SKUS + 1):
        cat = random.choice(CATEGORIES)
        skus.append((f"SKU-{cats_short[cat]}-{i:05d}", cat))
    return skus


SKU_LIST = generate_skus()
COSTCO_SKUS = set(random.sample([s[0] for s in SKU_LIST], int(NUM_SKUS * 0.58)))
SAMS_SKUS = set(random.sample([s[0] for s in SKU_LIST], int(NUM_SKUS * 0.82)))


def get_stores_for_retailer(retailer: str):
    if retailer == "Walmart":
        return [f"WMT-{r:02d}-{s:03d}" for r in range(1, 5) for s in range(1, WALMART_STORES // 4 + 4)][:WALMART_STORES]
    if retailer == "Sam's Club":
        return [f"SAM-{r:02d}-{s:03d}" for r in range(1, 5) for s in range(1, SAMS_STORES // 4 + 4)][:SAMS_STORES]
    if retailer == "Costco":
        return [f"COST-{r:02d}-{s:03d}" for r in range(1, 5) for s in range(1, COSTCO_STORES // 4 + 3)][:COSTCO_STORES]
    return []


def get_region_for_store(store: str) -> str:
    if store.startswith("WMT") or store.startswith("SAM") or store.startswith("COST"):
        try:
            r = int(store.split("-")[1])
            return REGIONS[(r - 1) % 4]
        except (IndexError, ValueError):
            pass
    return random.choice(REGIONS)


# --- Sales Table ---
def generate_sales():
    start_date = datetime(2024, 1, 1)
    rows = []
    retailers = ["Walmart", "Sam's Club", "Costco"]

    walmart_skus = [(s, c) for s, c in SKU_LIST]
    sams_skus = [(s, c) for s, c in SKU_LIST if s in SAMS_SKUS]
    costco_skus = [(s, c) for s, c in SKU_LIST if s in COSTCO_SKUS]

    for _ in range(NUM_SALES_ROWS):
        retailer = random.choices(
            retailers,
            weights=[0.54, 0.28, 0.18],
            k=1,
        )[0]

        sku_id, category = random.choice(
            walmart_skus if retailer == "Walmart" else (sams_skus if retailer == "Sam's Club" else costco_skus)
        )

        stores = get_stores_for_retailer(retailer)
        store = random.choice(stores)
        region = get_region_for_store(store)

        base_units = random.randint(1, 85)
        if retailer in ("Sam's Club", "Costco"):
            base_units = random.randint(2, 42)
        units = max(1, base_units)

        price = round(random.uniform(2.99, 62.99), 2)
        cost = round(price * random.uniform(0.48, 0.74), 2)
        revenue = round(units * price, 2)
        cost_total = round(units * cost, 2)
        margin_pct = round((revenue - cost_total) / revenue * 100, 2) if revenue else 0

        rows.append({
            "Date": (start_date + timedelta(days=random.randint(0, DAYS_OF_DATA - 1))).strftime("%Y-%m-%d"),
            "SKU": sku_id,
            "Product Category": category,
            "Retailer": retailer,
            "Store": store,
            "Region": region,
            "Units Sold": units,
            "Revenue": revenue,
            "Cost": cost_total,
            "Gross Margin %": margin_pct,
        })

    return rows


# --- Inventory Table: SKU, Store, Beginning Inventory, Ending Inventory, Reorder Point, Lead Time (days) ---
def generate_inventory():
    rows = []
    retailers = ["Walmart", "Sam's Club", "Costco"]

    for sku_id, category in SKU_LIST:
        for retailer in retailers:
            if retailer == "Costco" and sku_id not in COSTCO_SKUS:
                continue
            if retailer == "Sam's Club" and sku_id not in SAMS_SKUS:
                continue

            stores = get_stores_for_retailer(retailer)
            for store in random.sample(stores, min(len(stores), len(stores) // 2 + 6)):
                beg = random.randint(50, 950)
                ending = beg + random.randint(-280, 200)
                reorder = random.randint(20, 170)
                lead_time = random.randint(3, 14)
                region = get_region_for_store(store)

                rows.append({
                    "SKU": sku_id,
                    "Retailer": retailer,
                    "Store": store,
                    "Region": region,
                    "Beginning Inventory": max(0, beg),
                    "Ending Inventory": max(0, ending),
                    "Reorder Point": reorder,
                    "Lead Time (days)": lead_time,
                })
    return rows


# --- Pricing & Promo Table: SKU, List Price, Promo Price, Promo Start, Promo End, Marketing Spend ---
def generate_pricing_promo():
    rows = []
    retailers = ["Walmart", "Sam's Club", "Costco"]

    for sku_id, category in SKU_LIST:
        for retailer in retailers:
            if retailer == "Costco" and sku_id not in COSTCO_SKUS:
                continue
            if retailer == "Sam's Club" and sku_id not in SAMS_SKUS:
                continue

            list_price = round(random.uniform(3.99, 55.99), 2)
            promo_chance = {"Walmart": 0.52, "Sam's Club": 0.76, "Costco": 0.12}[retailer]
            has_promo = random.random() < promo_chance

            if has_promo:
                promo_pct = random.uniform(0.05, 0.24)
                promo_price = round(list_price * (1 - promo_pct), 2)
                start = datetime(2024, random.randint(1, 10), random.randint(1, 28))
                end = start + timedelta(days=random.randint(7, 45))
                marketing = round(random.uniform(400, 8500), 2)
            else:
                promo_price = list_price
                start = datetime(2024, 1, 1)
                end = datetime(2024, 12, 31)
                marketing = 0

            rows.append({
                "SKU": sku_id,
                "Retailer": retailer,
                "List Price": list_price,
                "Promo Price": promo_price,
                "Promo Start": start.strftime("%Y-%m-%d"),
                "Promo End": end.strftime("%Y-%m-%d"),
                "Marketing Spend": marketing,
            })
    return rows


def write_csv(path, rows, fieldnames):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(rows)
    print(f"Wrote {len(rows):,} rows to {path}")


def main():
    print("Generating Retail Supplier Performance data (80k-120k rows)...")
    sales = generate_sales()
    inv = generate_inventory()
    pricing = generate_pricing_promo()

    write_csv(
        DATA_DIR / "Sales.csv",
        sales,
        ["Date", "SKU", "Product Category", "Retailer", "Store", "Region", "Units Sold", "Revenue", "Cost", "Gross Margin %"],
    )
    write_csv(
        DATA_DIR / "Inventory.csv",
        inv,
        ["SKU", "Retailer", "Store", "Region", "Beginning Inventory", "Ending Inventory", "Reorder Point", "Lead Time (days)"],
    )
    write_csv(
        DATA_DIR / "Pricing_Promo.csv",
        pricing,
        ["SKU", "Retailer", "List Price", "Promo Price", "Promo Start", "Promo End", "Marketing Spend"],
    )
    print(f"Total Sales rows: {len(sales):,} (target 80k-120k)")
    print("Done.")


if __name__ == "__main__":
    main()
