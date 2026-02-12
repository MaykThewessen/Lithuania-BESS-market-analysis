"""
Fetch extended balancing/reserve data from ENTSO-E for Lithuania
================================================================
Covers the gap after Sep 2024 with:
- aFRR contracted reserve prices & quantities (A47)
- mFRR contracted reserve prices & quantities (A51)
- Activated balancing energy prices
- Imbalance volumes
"""

import os
import pandas as pd
import numpy as np
from entsoe import EntsoePandasClient
from dotenv import load_dotenv
import time
import warnings
warnings.filterwarnings('ignore')

load_dotenv(os.path.expanduser("~/.env"))
client = EntsoePandasClient(api_key=os.getenv("ENTSOE_API_KEY"))

TZ = 'Europe/Vilnius'
DATA_DIR = "/Users/mayk/LithuaniaBESS/data"

def safe_query(name, func, **kwargs):
    print(f"  Fetching {name}...", end=" ", flush=True)
    for attempt in range(3):
        try:
            data = func(**kwargs)
            if data is not None and len(data) > 0:
                print(f"OK ({len(data)} rows)")
                return data
            print("empty")
            return None
        except Exception as e:
            if attempt < 2:
                print(f"retry...", end=" ", flush=True)
                time.sleep(2)
            else:
                print(f"FAILED: {str(e)[:80]}")
                return None

# ============================================================
# 1. aFRR Reserve Prices & Quantities (process_type=A47)
# ============================================================
print("=== 1. aFRR Contracted Reserve Prices (A47) ===")
afrr_chunks = []
periods = [
    ('2024-06-01', '2024-10-01'),
    ('2024-10-01', '2025-01-01'),
    ('2025-01-01', '2025-04-01'),
    ('2025-04-01', '2025-07-01'),
    ('2025-07-01', '2025-10-01'),
    ('2025-10-01', '2026-01-01'),
    ('2026-01-01', '2026-02-13'),
]
for start_str, end_str in periods:
    chunk = safe_query(
        f"aFRR {start_str}",
        client.query_contracted_reserve_prices_procured_capacity,
        country_code='LT', process_type='A47',
        type_marketagreement_type='A01',
        start=pd.Timestamp(start_str, tz=TZ),
        end=pd.Timestamp(end_str, tz=TZ)
    )
    if chunk is not None:
        afrr_chunks.append(chunk)
    time.sleep(1)

if afrr_chunks:
    afrr_all = pd.concat(afrr_chunks)
    afrr_all = afrr_all[~afrr_all.index.duplicated(keep='first')]
    afrr_all.sort_index(inplace=True)
    afrr_all.to_csv(f"{DATA_DIR}/afrr_reserve_prices_LT.csv")
    print(f"  Saved: {len(afrr_all)} rows, {afrr_all.index.min()} to {afrr_all.index.max()}")
else:
    afrr_all = pd.DataFrame()

# ============================================================
# 2. mFRR Reserve Prices & Quantities (process_type=A51)
# ============================================================
print("\n=== 2. mFRR Contracted Reserve Prices (A51) ===")
mfrr_chunks = []
for start_str, end_str in periods:
    chunk = safe_query(
        f"mFRR {start_str}",
        client.query_contracted_reserve_prices_procured_capacity,
        country_code='LT', process_type='A51',
        type_marketagreement_type='A01',
        start=pd.Timestamp(start_str, tz=TZ),
        end=pd.Timestamp(end_str, tz=TZ)
    )
    if chunk is not None:
        mfrr_chunks.append(chunk)
    time.sleep(1)

if mfrr_chunks:
    mfrr_all = pd.concat(mfrr_chunks)
    mfrr_all = mfrr_all[~mfrr_all.index.duplicated(keep='first')]
    mfrr_all.sort_index(inplace=True)
    mfrr_all.to_csv(f"{DATA_DIR}/mfrr_reserve_prices_LT.csv")
    print(f"  Saved: {len(mfrr_all)} rows, {mfrr_all.index.min()} to {mfrr_all.index.max()}")
else:
    mfrr_all = pd.DataFrame()

# ============================================================
# 3. Activated Balancing Energy Prices
# ============================================================
print("\n=== 3. Activated Balancing Energy Prices ===")
act_chunks = []
for start_str, end_str in periods:
    chunk = safe_query(
        f"activated energy prices {start_str}",
        client.query_activated_balancing_energy_prices,
        country_code='LT',
        start=pd.Timestamp(start_str, tz=TZ),
        end=pd.Timestamp(end_str, tz=TZ)
    )
    if chunk is not None:
        act_chunks.append(chunk)
    time.sleep(1)

if act_chunks:
    act_all = pd.concat(act_chunks)
    act_all = act_all[~act_all.index.duplicated(keep='first')]
    act_all.sort_index(inplace=True)
    act_all.to_csv(f"{DATA_DIR}/activated_balancing_energy_prices_LT.csv")
    print(f"  Saved: {len(act_all)} rows, {act_all.index.min()} to {act_all.index.max()}")
else:
    act_all = pd.DataFrame()

# ============================================================
# 4. Imbalance Volumes (extended)
# ============================================================
print("\n=== 4. Imbalance Volumes (extended) ===")
imb_vol_chunks = []
for start_str, end_str in [
    ('2024-09-01', '2024-11-01'),
    ('2024-11-01', '2025-01-01'),
    ('2025-01-01', '2025-04-01'),
    ('2025-04-01', '2025-07-01'),
    ('2025-07-01', '2025-10-01'),
    ('2025-10-01', '2026-01-01'),
    ('2026-01-01', '2026-02-13'),
]:
    chunk = safe_query(
        f"imbalance volumes {start_str}",
        client.query_imbalance_volumes,
        country_code='LT',
        start=pd.Timestamp(start_str, tz=TZ),
        end=pd.Timestamp(end_str, tz=TZ)
    )
    if chunk is not None:
        imb_vol_chunks.append(chunk)
    time.sleep(1)

if imb_vol_chunks:
    imb_vol_all = pd.concat(imb_vol_chunks)
    imb_vol_all = imb_vol_all[~imb_vol_all.index.duplicated(keep='first')]
    imb_vol_all.sort_index(inplace=True)
    imb_vol_all.to_csv(f"{DATA_DIR}/imbalance_volumes_extended_LT.csv")
    print(f"  Saved: {len(imb_vol_all)} rows, {imb_vol_all.index.min()} to {imb_vol_all.index.max()}")
else:
    imb_vol_all = pd.DataFrame()

# ============================================================
# Summary statistics
# ============================================================
print("\n" + "="*60)
print("SUMMARY OF EXTENDED BALANCING DATA")
print("="*60)

if len(afrr_all) > 0:
    print("\naFRR Reserve (A47):")
    for col in afrr_all.columns:
        vals = pd.to_numeric(afrr_all[col], errors='coerce')
        # Annual summary
        annual = vals.groupby(vals.index.year).agg(['mean', 'median', 'min', 'max', 'count'])
        print(f"\n  {col}:")
        for yr, row in annual.iterrows():
            print(f"    {yr}: mean={row['mean']:.2f}, median={row['median']:.2f}, "
                  f"min={row['min']:.2f}, max={row['max']:.2f}, count={int(row['count'])}")

if len(mfrr_all) > 0:
    print("\nmFRR Reserve (A51):")
    for col in mfrr_all.columns:
        vals = pd.to_numeric(mfrr_all[col], errors='coerce')
        annual = vals.groupby(vals.index.year).agg(['mean', 'median', 'min', 'max', 'count'])
        print(f"\n  {col}:")
        for yr, row in annual.iterrows():
            print(f"    {yr}: mean={row['mean']:.2f}, median={row['median']:.2f}, "
                  f"min={row['min']:.2f}, max={row['max']:.2f}, count={int(row['count'])}")

if len(act_all) > 0:
    print("\nActivated Balancing Energy Prices:")
    print(f"  Total activations: {len(act_all)}")
    if 'Direction' in act_all.columns:
        for direction in act_all['Direction'].unique():
            mask = act_all['Direction'] == direction
            prices = pd.to_numeric(act_all.loc[mask, 'Price'], errors='coerce')
            print(f"  {direction}: count={mask.sum()}, mean={prices.mean():.2f}, "
                  f"median={prices.median():.2f}, max={prices.max():.2f}")

print("\nDone! All extended balancing data saved.")
