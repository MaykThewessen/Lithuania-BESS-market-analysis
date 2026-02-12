"""
Fetch Lithuania electricity market data from ENTSO-E Transparency Platform
==========================================================================
Retrieves: day-ahead prices, imbalance prices, load, generation by type,
           cross-border flows, installed capacity
"""

import os
import pandas as pd
import numpy as np
from entsoe import EntsoePandasClient
from dotenv import load_dotenv
import json
import time
import warnings
warnings.filterwarnings('ignore')

# Load API key
load_dotenv(os.path.expanduser("~/.env"))
API_KEY = os.getenv("ENTSOE_API_KEY")

if not API_KEY:
    raise ValueError("ENTSOE_API_KEY not found in ~/.env")

client = EntsoePandasClient(api_key=API_KEY)

COUNTRY = 'LT'
TZ = 'Europe/Vilnius'
OUTPUT_DIR = "/Users/mayk/LithuaniaBESS/data"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Date ranges - fetch 2021 through now
START = pd.Timestamp('2021-01-01', tz=TZ)
END = pd.Timestamp('2026-02-12', tz=TZ)

results = {}

def safe_query(name, func, *args, **kwargs):
    """Run an ENTSO-E query with error handling and retry."""
    print(f"  Fetching {name}...", end=" ", flush=True)
    for attempt in range(3):
        try:
            data = func(*args, **kwargs)
            print(f"OK ({len(data)} rows)" if hasattr(data, '__len__') else "OK")
            return data
        except Exception as e:
            if attempt < 2:
                print(f"retry {attempt+1}...", end=" ", flush=True)
                time.sleep(2)
            else:
                print(f"FAILED: {e}")
                return None

# ============================================================
# 1. Day-Ahead Prices
# ============================================================
print("\n=== 1. Day-Ahead Prices ===")
da_prices = safe_query(
    "day-ahead prices",
    client.query_day_ahead_prices, COUNTRY, start=START, end=END
)
if da_prices is not None:
    results['da_prices'] = da_prices
    da_prices.to_csv(f"{OUTPUT_DIR}/da_prices_LT.csv")

# ============================================================
# 2. Imbalance Prices
# ============================================================
print("\n=== 2. Imbalance Prices ===")
# Try fetching in yearly chunks to avoid timeouts
imbalance_all = []
for year in range(2021, 2027):
    y_start = pd.Timestamp(f'{year}-01-01', tz=TZ)
    y_end = min(pd.Timestamp(f'{year+1}-01-01', tz=TZ), END)
    if y_start >= END:
        break
    chunk = safe_query(
        f"imbalance prices {year}",
        client.query_imbalance_prices, COUNTRY, start=y_start, end=y_end
    )
    if chunk is not None:
        imbalance_all.append(chunk)
    time.sleep(1)

if imbalance_all:
    imb_prices = pd.concat(imbalance_all)
    results['imbalance_prices'] = imb_prices
    imb_prices.to_csv(f"{OUTPUT_DIR}/imbalance_prices_LT.csv")

# ============================================================
# 3. Actual Total Load
# ============================================================
print("\n=== 3. Actual Total Load ===")
load_all = []
for year in range(2021, 2027):
    y_start = pd.Timestamp(f'{year}-01-01', tz=TZ)
    y_end = min(pd.Timestamp(f'{year+1}-01-01', tz=TZ), END)
    if y_start >= END:
        break
    chunk = safe_query(
        f"actual load {year}",
        client.query_load, COUNTRY, start=y_start, end=y_end
    )
    if chunk is not None:
        load_all.append(chunk)
    time.sleep(1)

if load_all:
    load_data = pd.concat(load_all)
    results['load'] = load_data
    load_data.to_csv(f"{OUTPUT_DIR}/actual_load_LT.csv")

# ============================================================
# 4. Generation by Type (Actual Aggregated)
# ============================================================
print("\n=== 4. Generation by Type ===")
gen_all = []
for year in range(2021, 2027):
    y_start = pd.Timestamp(f'{year}-01-01', tz=TZ)
    y_end = min(pd.Timestamp(f'{year+1}-01-01', tz=TZ), END)
    if y_start >= END:
        break
    chunk = safe_query(
        f"generation by type {year}",
        client.query_generation, COUNTRY, start=y_start, end=y_end
    )
    if chunk is not None:
        gen_all.append(chunk)
    time.sleep(1)

if gen_all:
    gen_data = pd.concat(gen_all)
    results['generation'] = gen_data
    gen_data.to_csv(f"{OUTPUT_DIR}/generation_by_type_LT.csv")

# ============================================================
# 5. Installed Generation Capacity
# ============================================================
print("\n=== 5. Installed Generation Capacity ===")
for year in range(2021, 2027):
    y_start = pd.Timestamp(f'{year}-01-01', tz=TZ)
    y_end = pd.Timestamp(f'{year}-12-31', tz=TZ)
    if y_start >= END:
        break
    cap = safe_query(
        f"installed capacity {year}",
        client.query_installed_generation_capacity, COUNTRY, start=y_start, end=y_end
    )
    if cap is not None:
        results[f'capacity_{year}'] = cap
        cap.to_csv(f"{OUTPUT_DIR}/installed_capacity_LT_{year}.csv")
    time.sleep(1)

# ============================================================
# 6. Cross-border Flows (key interconnections)
# ============================================================
print("\n=== 6. Cross-border Flows ===")
neighbors = {
    'SE_4': 'Sweden (NordBalt)',
    'PL': 'Poland (LitPol)',
    'LV': 'Latvia',
}

for neighbor_code, neighbor_name in neighbors.items():
    print(f"\n  --- {neighbor_name} ---")
    # Import to LT
    imp = safe_query(
        f"flow {neighbor_code}->LT",
        client.query_crossborder_flows, neighbor_code, COUNTRY, start=START, end=END
    )
    if imp is not None:
        results[f'flow_{neighbor_code}_to_LT'] = imp
        imp.to_csv(f"{OUTPUT_DIR}/flow_{neighbor_code}_to_LT.csv")
    time.sleep(1)

    # Export from LT
    exp = safe_query(
        f"flow LT->{neighbor_code}",
        client.query_crossborder_flows, COUNTRY, neighbor_code, start=START, end=END
    )
    if exp is not None:
        results[f'flow_LT_to_{neighbor_code}'] = exp
        exp.to_csv(f"{OUTPUT_DIR}/flow_LT_to_{neighbor_code}.csv")
    time.sleep(1)

# ============================================================
# Summary
# ============================================================
print("\n" + "="*60)
print("DATA RETRIEVAL SUMMARY")
print("="*60)
for key, val in results.items():
    if isinstance(val, (pd.Series, pd.DataFrame)):
        shape = val.shape
        idx_start = val.index.min() if len(val) > 0 else 'N/A'
        idx_end = val.index.max() if len(val) > 0 else 'N/A'
        print(f"  {key}: shape={shape}, range={idx_start} to {idx_end}")
    else:
        print(f"  {key}: {type(val)}")

print(f"\nAll CSV files saved to: {OUTPUT_DIR}/")
print("Done!")
