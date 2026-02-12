"""
Optimize HTML Report — Self-Contained with Base64 Embedded Charts
=================================================================
Renders all Plotly charts to SVG, embeds as base64 data URIs,
removes CDN dependency, inlines all CSS, compresses output.
"""

import os
import base64
import json
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime

DATA_DIR = "/Users/mayk/LithuaniaBESS/data"
OUT_DIR = "/Users/mayk/LithuaniaBESS"

# Chart rendering config
CHART_WIDTH = 1100
CHART_HEIGHT = 480
CHART_FONT = dict(family="Calibri, Arial, sans-serif", size=13)
CHART_COLORS = {
    'primary': '#1F4E79', 'accent': '#2E75B6', 'green': '#70AD47',
    'solar': '#FFC000', 'red': '#C00000', 'purple': '#7B2D8E',
    'orange': '#ED7D31', 'gray': '#A5A5A5'
}
RT_EFF = 0.88
DA_CAPTURE = 0.85
AFRR_AVAIL = {1: 0.65, 2: 0.80, 4: 0.90}
FCR_AVAIL = {1: 0.90, 2: 0.92, 4: 0.95}
MFRR_AVAIL = {1: 0.70, 2: 0.82, 4: 0.90}
FCR_PRICE_PER_HOUR = {2024: 0, 2025: 30, 2026: 22, 2027: 18, 2028: 15, 2029: 12, 2030: 10}
MULTI_MARKET_ALLOC = {'aFRR': 0.40, 'FCR': 0.20, 'mFRR': 0.05, 'DA': 0.25, 'Imbalance': 0.10}
COMPRESSION = {2025: 1.0, 2026: 0.65, 2027: 0.45, 2028: 0.35, 2029: 0.30, 2030: 0.28}
DA_COMPRESSION = {2025: 1.0, 2026: 0.85, 2027: 0.70, 2028: 0.60, 2029: 0.55, 2030: 0.50}
PERIODS_PER_YEAR = 35_040


def fig_to_base64(fig, width=CHART_WIDTH, height=CHART_HEIGHT):
    """Render Plotly figure to base64 SVG data URI."""
    svg_bytes = fig.to_image(format="svg", width=width, height=height)
    b64 = base64.b64encode(svg_bytes).decode('utf-8')
    return f"data:image/svg+xml;base64,{b64}"


def img_tag(b64_uri, alt="Chart", style="width:100%; max-width:1100px;"):
    return f'<img src="{b64_uri}" alt="{alt}" style="{style}" loading="lazy">'


# ============================================================
# Load all data
# ============================================================
print("Loading data...")

da = pd.read_csv(f"{DATA_DIR}/da_prices_LT.csv", index_col=0, parse_dates=True)
da.columns = ['price']
da.index = pd.to_datetime(da.index, utc=True)
da['price'] = pd.to_numeric(da['price'], errors='coerce')

afrr = pd.read_csv(f"{DATA_DIR}/afrr_reserve_prices_LT.csv", index_col=0, parse_dates=True)
afrr.index = pd.to_datetime(afrr.index, utc=True)
for c in afrr.columns: afrr[c] = pd.to_numeric(afrr[c], errors='coerce')

mfrr = pd.read_csv(f"{DATA_DIR}/mfrr_reserve_prices_LT.csv", index_col=0, parse_dates=True)
mfrr.index = pd.to_datetime(mfrr.index, utc=True)
for c in mfrr.columns: mfrr[c] = pd.to_numeric(mfrr[c], errors='coerce')

imb = pd.read_csv(f"{DATA_DIR}/imbalance_prices_LT.csv", index_col=0, parse_dates=True)
imb.index = pd.to_datetime(imb.index, utc=True)
for c in imb.columns: imb[c] = pd.to_numeric(imb[c], errors='coerce')

load_data = pd.read_csv(f"{DATA_DIR}/actual_load_LT.csv", index_col=0, parse_dates=True)
load_data.index = pd.to_datetime(load_data.index, utc=True)
load_col = load_data.columns[0]
load_data[load_col] = pd.to_numeric(load_data[load_col], errors='coerce')

gen = pd.read_csv(f"{DATA_DIR}/generation_by_type_LT.csv", index_col=0, parse_dates=True)
gen.index = pd.to_datetime(gen.index, utc=True)
for c in gen.columns: gen[c] = pd.to_numeric(gen[c], errors='coerce')

print(f"  DA: {len(da)}, aFRR: {len(afrr)}, mFRR: {len(mfrr)}, Imb: {len(imb)}, Load: {len(load_data)}, Gen: {len(gen)}")

# ============================================================
# Render all charts
# ============================================================
charts = {}
layout_defaults = dict(
    font=CHART_FONT,
    plot_bgcolor='white', paper_bgcolor='white',
    margin=dict(l=60, r=30, t=50, b=60),
    legend=dict(orientation='h', y=-0.18),
)

# --- 1. DA Monthly Prices ---
print("Rendering charts...")
print("  1. DA Monthly Prices")
fig = go.Figure()
for year in sorted(da.index.year.unique()):
    if year < 2021: continue
    mask = da.index.year == year
    monthly = da.loc[mask, 'price'].resample('M').mean()
    fig.add_trace(go.Scatter(
        x=[f"{m.month:02d}" for m in monthly.index],
        y=monthly.values, name=str(year), mode='lines+markers',
        line=dict(width=2.5), marker=dict(size=5)
    ))
fig.update_layout(title='Monthly Average Day-Ahead Price by Year',
                  yaxis_title='EUR/MWh', xaxis_title='Month',
                  xaxis=dict(ticktext=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'],
                             tickvals=[f"{i:02d}" for i in range(1,13)]),
                  **layout_defaults)
charts['da_monthly'] = fig_to_base64(fig)

# --- 2. Daily Price Spread ---
print("  2. Daily Spread")
fig = go.Figure()
for year in range(2021, 2027):
    mask = da.index.year == year
    if mask.sum() == 0: continue
    daily = da.loc[mask, 'price'].groupby(da.loc[mask].index.date)
    spreads = daily.max() - daily.min()
    fig.add_trace(go.Box(y=spreads.values, name=str(year),
                         marker_color=CHART_COLORS['accent']))
fig.update_layout(title='Daily Price Spread (Max - Min) by Year — Arbitrage Potential',
                  yaxis_title='EUR/MWh', showlegend=False, **layout_defaults)
charts['spread'] = fig_to_base64(fig)

# --- 3. Hourly Profile ---
print("  3. Hourly Profile")
fig = go.Figure()
for year in [2023, 2024, 2025]:
    mask = da.index.year == year
    if mask.sum() == 0: continue
    hourly = da.loc[mask, 'price'].groupby(da.loc[mask].index.hour).mean()
    fig.add_trace(go.Scatter(x=list(range(24)), y=hourly.values,
                             name=str(year), mode='lines+markers',
                             line=dict(width=2.5), marker=dict(size=5)))
fig.update_layout(title='Average Hourly Price Profile',
                  yaxis_title='EUR/MWh', xaxis_title='Hour of Day',
                  xaxis=dict(dtick=2), **layout_defaults)
charts['hourly'] = fig_to_base64(fig)

# --- 4. Negative Price Hours ---
print("  4. Negative Hours")
neg_hours = {}
for year in range(2021, 2027):
    mask = da.index.year == year
    if mask.sum() > 0:
        neg_hours[year] = int((da.loc[mask, 'price'] < 0).sum())
fig = go.Figure(data=[go.Bar(
    x=list(neg_hours.keys()), y=list(neg_hours.values()),
    marker_color=[CHART_COLORS['red'] if v > 100 else CHART_COLORS['accent'] for v in neg_hours.values()],
    text=list(neg_hours.values()), textposition='auto'
)])
fig.update_layout(title='Negative Price Hours per Year',
                  yaxis_title='Hours', showlegend=False, **layout_defaults)
charts['neg_hours'] = fig_to_base64(fig)

# --- 5. Imbalance Prices ---
print("  5. Imbalance Prices")
fig = go.Figure()
for year in range(2021, 2025):
    mask = imb.index.year == year
    if mask.sum() == 0: continue
    fig.add_trace(go.Bar(x=[year], y=[imb.loc[mask, 'Long'].mean()], name='Avg Long' if year == 2021 else None,
                         marker_color=CHART_COLORS['green'], showlegend=(year==2021)))
    fig.add_trace(go.Bar(x=[year], y=[imb.loc[mask, 'Short'].mean()], name='Avg Short' if year == 2021 else None,
                         marker_color=CHART_COLORS['red'], showlegend=(year==2021)))
fig.update_layout(title='Annual Imbalance Prices (EUR/MWh)',
                  yaxis_title='EUR/MWh', barmode='group', **layout_defaults)
charts['imbalance'] = fig_to_base64(fig)

# --- 6. Load Annual ---
print("  6. Load")
load_annual = {}
for year in range(2021, 2027):
    mask = load_data.index.year == year
    if mask.sum() == 0: continue
    v = load_data.loc[mask, load_col]
    load_annual[year] = {'avg': v.mean(), 'max': v.max(), 'min': v.min(), 'twh': v.sum() / 1e6 * (len(v) / (8760 * (4 if len(load_data) > 50000 else 1)))}

years_load = sorted(load_annual.keys())
fig = go.Figure()
fig.add_trace(go.Bar(x=years_load, y=[load_annual[y]['avg']/1000 for y in years_load],
                     name='Avg Load (GW)', marker_color=CHART_COLORS['accent']))
fig.add_trace(go.Scatter(x=years_load, y=[load_annual[y]['max']/1000 for y in years_load],
                         name='Peak Load (GW)', mode='lines+markers',
                         marker=dict(color=CHART_COLORS['red'], size=8), line=dict(width=2.5)))
fig.add_trace(go.Scatter(x=years_load, y=[load_annual[y]['min']/1000 for y in years_load],
                         name='Min Load (GW)', mode='lines+markers',
                         marker=dict(color=CHART_COLORS['green'], size=8), line=dict(width=2.5, dash='dot')))
fig.update_layout(title='Annual Electricity Load (GW)',
                  yaxis_title='GW', **layout_defaults)
charts['load'] = fig_to_base64(fig)

# --- 7. Generation by Type ---
print("  7. Generation Mix")
type_map = {
    'Wind Onshore': 'Wind', 'Solar': 'Solar', 'Fossil Gas': 'Gas',
    'Hydro Run-of-river and poundage': 'Hydro', 'Biomass': 'Biomass',
    'Energy storage': 'Battery', 'Waste': 'Waste'
}
type_colors = {'Wind': CHART_COLORS['accent'], 'Solar': CHART_COLORS['solar'],
               'Gas': CHART_COLORS['gray'], 'Hydro': '#00B0F0',
               'Biomass': CHART_COLORS['green'], 'Battery': CHART_COLORS['orange'], 'Waste': '#999'}

# Determine resolution
is_15min = len(gen) > len(da) * 1.5
hours_per_row = 0.25 if is_15min else 1.0
gen_twh = {}
for year in range(2021, 2027):
    mask = gen.index.year == year
    if mask.sum() == 0: continue
    gen_twh[year] = {}
    for raw_col, nice_name in type_map.items():
        if raw_col in gen.columns:
            val = gen.loc[mask, raw_col].sum() * hours_per_row / 1e6
            if val > 0.001:
                gen_twh[year][nice_name] = round(val, 3)

fig = go.Figure()
all_types = sorted(set(t for y in gen_twh.values() for t in y.keys()))
for t in all_types:
    fig.add_trace(go.Bar(
        x=list(gen_twh.keys()),
        y=[gen_twh[y].get(t, 0) for y in gen_twh.keys()],
        name=t, marker_color=type_colors.get(t, '#999')
    ))
fig.update_layout(title='Generation by Source (TWh/yr)',
                  yaxis_title='TWh', barmode='stack', **layout_defaults)
charts['gen_stack'] = fig_to_base64(fig)

# --- 8. Cross-border Flows ---
print("  8. Cross-border Flows")
neighbors = {'SE_4': 'Sweden (NordBalt)', 'PL': 'Poland (LitPol)', 'LV': 'Latvia'}
flow_annual = {}
for code, name in neighbors.items():
    for direction, label in [('to_LT', 'Import'), ('LT_to', 'Export')]:
        fpath = f"{DATA_DIR}/flow_{direction.replace('to_LT', code + '_to_LT').replace('LT_to', 'LT_to_' + code)}.csv"
        if not os.path.exists(fpath): continue
        fdata = pd.read_csv(fpath, index_col=0, parse_dates=True)
        fdata.index = pd.to_datetime(fdata.index, utc=True)
        col = fdata.columns[0]
        fdata[col] = pd.to_numeric(fdata[col], errors='coerce')
        for year in range(2021, 2027):
            mask = fdata.index.year == year
            if mask.sum() == 0: continue
            key = f"{name} {label}"
            if key not in flow_annual: flow_annual[key] = {}
            flow_annual[key][year] = round(fdata.loc[mask, col].mean(), 1)

fig = go.Figure()
for key, data in flow_annual.items():
    fig.add_trace(go.Bar(x=list(data.keys()), y=list(data.values()), name=key))
fig.update_layout(title='Cross-Border Flows Annual Average (MW)',
                  yaxis_title='MW', barmode='group', **layout_defaults)
charts['flows'] = fig_to_base64(fig)

# --- 9. Installed Capacity ---
print("  9. Installed Capacity")
years_cap = list(range(2020, 2031))
wind_cumul = [668, 668, 946, 920, 1070, 1200, 1400, 1800, 2500, 3200, 4000]
pv_cumul = [100, 250, 520, 1410, 2080, 2800, 3500, 4000, 4500, 5000, 5100]
bess_power = [0, 200, 200, 250, 535, 800, 1200, 1500, 1700, 1800, 1900]
bess_energy = [0, 200, 200, 350, 1070, 1600, 2800, 3500, 4000, 4200, 4500]
fossil_mw = [3100, 3100, 3100, 3100, 3100, 3100, 3000, 2800, 2600, 2400, 2200]

fig = go.Figure()
fig.add_trace(go.Scatter(x=years_cap, y=wind_cumul, name='Wind', mode='lines+markers',
                         line=dict(color=CHART_COLORS['accent'], width=2.5)))
fig.add_trace(go.Scatter(x=years_cap, y=pv_cumul, name='Solar PV', mode='lines+markers',
                         line=dict(color=CHART_COLORS['solar'], width=2.5)))
fig.add_trace(go.Bar(x=years_cap, y=bess_power, name='BESS (MW)', marker_color=CHART_COLORS['orange'], opacity=0.7))
fig.add_trace(go.Scatter(x=years_cap, y=fossil_mw, name='Fossil', mode='lines+markers',
                         line=dict(color=CHART_COLORS['gray'], width=2, dash='dash')))
fig.add_vrect(x0=2025.5, x1=2030.5, fillcolor='rgba(0,0,0,0.03)', line_width=0,
              annotation_text='Forecast', annotation_position='top left')
fig.update_layout(title='Installed Generation Capacity (MW)',
                  yaxis_title='MW', **layout_defaults)
charts['capacity'] = fig_to_base64(fig)

# --- 10. aFRR Monthly ---
print("  10. aFRR Monthly")
afrr_mkt = afrr[afrr.index >= '2024-10-01']
afrr_monthly = afrr_mkt.resample('M').agg({'Up Prices': ['mean', 'median'],
                                            'Down Prices': ['mean', 'median'],
                                            'Up Quantity': 'mean'})
afrr_monthly.columns = ['up_mean', 'up_median', 'down_mean', 'down_median', 'up_qty']
fig = make_subplots(specs=[[{"secondary_y": True}]])
months_str = [m.strftime('%Y-%m') for m in afrr_monthly.index]
fig.add_trace(go.Bar(x=months_str, y=afrr_monthly['up_qty'], name='Up Quantity (MW)',
                     marker_color='rgba(46,117,182,0.3)'), secondary_y=True)
fig.add_trace(go.Scatter(x=months_str, y=afrr_monthly['up_mean'], name='Up Price Mean',
                         line=dict(color=CHART_COLORS['red'], width=2.5), mode='lines+markers'))
fig.add_trace(go.Scatter(x=months_str, y=afrr_monthly['up_median'], name='Up Price Median',
                         line=dict(color=CHART_COLORS['orange'], width=2, dash='dot'), mode='lines'))
fig.add_trace(go.Scatter(x=months_str, y=afrr_monthly['down_mean'], name='Down Price Mean',
                         line=dict(color=CHART_COLORS['green'], width=2.5), mode='lines+markers'))
fig.update_layout(title='aFRR Contracted Reserve Prices — Monthly (Post-PICASSO)',
                  **layout_defaults)
fig.update_yaxes(title_text='EUR/MW per ISP', secondary_y=False)
fig.update_yaxes(title_text='MW', secondary_y=True)
charts['afrr_monthly'] = fig_to_base64(fig)

# --- 11. Revenue by Market 2025 ---
print("  11. Revenue by Market 2025")

# Compute revenues (same logic as add_revenue_section.py)
def compute_da_rev(year, dur):
    mask = da.index.year == year
    if mask.sum() == 0: return 0
    yearly = da.loc[mask, 'price'].dropna()
    daily = yearly.groupby(yearly.index.date)
    total = 0; days = 0
    for _, grp in daily:
        if len(grp) < 24: continue
        p = np.sort(grp.values)
        rev = p[-dur:].sum() * np.sqrt(RT_EFF) - p[:dur].sum() / np.sqrt(RT_EFF)
        total += max(0, rev); days += 1
    return (total / days * 365 * DA_CAPTURE) if days > 0 else 0

def compute_afrr_rev(year, dur):
    mask = afrr.index.year == year
    if mask.sum() == 0: return 0
    yr = afrr.loc[mask]
    ann = PERIODS_PER_YEAR / len(yr)
    return (0.5 * yr['Up Prices'].sum() + 0.5 * yr['Down Prices'].sum()) * AFRR_AVAIL[dur] * ann

def compute_fcr_rev(year, dur):
    price = FCR_PRICE_PER_HOUR.get(year, 15)
    hours = 8760 * (11/12) if year == 2025 else 8760
    return price * hours * FCR_AVAIL[dur]

def compute_mfrr_rev(year, dur):
    mask = mfrr.index.year == year
    if mask.sum() == 0: return 0
    yr = mfrr.loc[mask]
    ann = PERIODS_PER_YEAR / len(yr)
    return (0.5 * yr['Up Prices'].sum() + 0.5 * yr['Down Prices'].sum()) * MFRR_AVAIL.get(dur, 0.7) * ann

def compute_imb_rev(year, dur):
    mask_da = da.index.year == year
    mask_imb = imb.index.year == year
    if mask_da.sum() == 0 or mask_imb.sum() == 0: return 0
    common = da.loc[mask_da].index.intersection(imb.loc[mask_imb].index)
    if len(common) == 0: return 0
    spread = (imb.loc[common, 'Short'] - da.loc[common, 'price']).abs()
    daily = spread.groupby(spread.index.date)
    total = 0; days = 0
    for _, grp in daily:
        total += grp.nlargest(dur).sum() * RT_EFF; days += 1
    return (total / days * 365 * DA_CAPTURE) if days > 0 else 0

durations = [1, 2, 4]
markets = ['DA Arbitrage', 'aFRR', 'FCR', 'mFRR', 'Imbalance']
rev_2025 = {}
rev_2024 = {}
for dur in durations:
    for year, store in [(2025, rev_2025), (2024, rev_2024)]:
        store[('DA Arbitrage', dur)] = compute_da_rev(year, dur)
        store[('aFRR', dur)] = compute_afrr_rev(year, dur)
        store[('FCR', dur)] = compute_fcr_rev(year, dur)
        store[('mFRR', dur)] = compute_mfrr_rev(year, dur)
        store[('Imbalance', dur)] = compute_imb_rev(year, dur)
        combined = sum(
            MULTI_MARKET_ALLOC[m] * store[(m, dur)] / {
                'DA': DA_CAPTURE, 'aFRR': AFRR_AVAIL[dur], 'FCR': FCR_AVAIL[dur],
                'mFRR': MFRR_AVAIL[dur], 'Imbalance': DA_CAPTURE
            }.get(m, 1) for m in ['DA', 'aFRR', 'FCR', 'mFRR', 'Imbalance']
            if (m, dur) in store
        )
        store[('Combined', dur)] = combined

dur_labels = ['1h BESS', '2h BESS', '4h BESS']
fig = go.Figure()
for mkt in markets:
    fig.add_trace(go.Bar(
        x=dur_labels, y=[rev_2025.get((mkt, d), 0) for d in durations],
        name=mkt
    ))
fig.update_layout(title='BESS Revenue by Market Segment — 2025 Actual (EUR/MW/year)',
                  yaxis_title='EUR/MW/year', yaxis_tickformat=',',
                  barmode='stack', **layout_defaults)
charts['rev_2025'] = fig_to_base64(fig)

# --- 12. Revenue Projection ---
print("  12. Revenue Projection")
proj = {}
for year in range(2025, 2031):
    comp = COMPRESSION.get(year, 0.28)
    da_comp = DA_COMPRESSION.get(year, 0.50)
    for dur in durations:
        base_da = rev_2025.get(('DA Arbitrage', dur), 0)
        base_afrr = rev_2025.get(('aFRR', dur), 0)
        base_mfrr = rev_2025.get(('mFRR', dur), 0)
        base_imb = rev_2024.get(('Imbalance', dur), 0)
        combined = (
            MULTI_MARKET_ALLOC['aFRR'] * base_afrr * comp / AFRR_AVAIL[dur] +
            MULTI_MARKET_ALLOC['FCR'] * compute_fcr_rev(year, dur) / FCR_AVAIL[dur] +
            MULTI_MARKET_ALLOC['mFRR'] * base_mfrr * comp / MFRR_AVAIL[dur] +
            MULTI_MARKET_ALLOC['DA'] * base_da * da_comp / DA_CAPTURE +
            MULTI_MARKET_ALLOC['Imbalance'] * base_imb * da_comp / DA_CAPTURE
        )
        proj[(year, dur)] = combined

fig = go.Figure()
years_proj = list(range(2025, 2031))
for dur in durations:
    fig.add_trace(go.Bar(
        x=[str(y) for y in years_proj],
        y=[round(proj.get((y, dur), 0)) for y in years_proj],
        name=f'{dur}h BESS'
    ))
fig.add_annotation(x='2025', y=max(proj.get((2025, d), 0) for d in durations),
                   text='Post-BRELL<br>scarcity peak', showarrow=True, arrowhead=2, ax=40, ay=-40,
                   font=dict(size=11, color='#c00'))
fig.add_annotation(x='2027', y=max(proj.get((2027, d), 0) for d in durations),
                   text='~1,200 MW BESS', showarrow=True, arrowhead=2, ax=40, ay=-40,
                   font=dict(size=11, color='#666'))
fig.update_layout(title='Multi-Market Combined Revenue Projection (EUR/MW/year)',
                  yaxis_title='EUR/MW/year', yaxis_tickformat=',',
                  barmode='group', **layout_defaults)
charts['rev_projection'] = fig_to_base64(fig)

# --- 13. Pipeline vs Market ---
print("  13. Pipeline vs Market")
afrr_up_mean = afrr.loc[afrr.index.year == 2025, 'Up Quantity'].mean()
afrr_down_mean = afrr.loc[afrr.index.year == 2025, 'Down Quantity'].mean()
mfrr_up_mean = mfrr.loc[mfrr.index.year == 2025, 'Up Quantity'].mean()
mfrr_down_mean = mfrr.loc[mfrr.index.year == 2025, 'Down Quantity'].mean()
total_bal = afrr_up_mean + afrr_down_mean + mfrr_up_mean + mfrr_down_mean + 40

segments = ['aFRR Up', 'aFRR Down', 'mFRR', 'FCR (est.)', 'Total Balancing']
seg_values = [round(afrr_up_mean), round(afrr_down_mean), round(mfrr_up_mean + mfrr_down_mean), 40, round(total_bal)]

fig = go.Figure()
fig.add_trace(go.Bar(x=segments, y=seg_values, name='Market Demand (MW)',
                     marker_color=[CHART_COLORS['accent']]*4 + [CHART_COLORS['primary']]))
fig.add_hline(y=1700, line_dash='dash', line_color=CHART_COLORS['red'], line_width=3,
              annotation_text='Pipeline: 1,700 MW', annotation_position='top right')
fig.add_hline(y=454, line_dash='dot', line_color=CHART_COLORS['green'], line_width=3,
              annotation_text='Installed: 454 MW', annotation_position='top right')
fig.update_layout(title='BESS Pipeline (1.7 GW) vs Actual Market Demand',
                  yaxis_title='MW', **layout_defaults)
charts['pipeline'] = fig_to_base64(fig)

# --- 14. Build-out Scenarios ---
print("  14. Build-out Scenarios")
scenarios = {
    'High': {2025: 454, 2026: 800, 2027: 1200, 2028: 1500, 2029: 1700, 2030: 1700},
    'Base': {2025: 454, 2026: 650, 2027: 950, 2028: 1200, 2029: 1400, 2030: 1500},
    'Low': {2025: 454, 2026: 550, 2027: 700, 2028: 850, 2029: 1000, 2030: 1100},
}
years_sc = list(range(2025, 2031))

fig = go.Figure()
sc_colors = {'High': CHART_COLORS['red'], 'Base': CHART_COLORS['solar'], 'Low': CHART_COLORS['green']}
for name, sc in scenarios.items():
    fig.add_trace(go.Scatter(x=years_sc, y=[sc[y] for y in years_sc], name=f'{name} Scenario',
                             mode='lines+markers', line=dict(color=sc_colors[name], width=3 if name == 'Base' else 2),
                             marker=dict(size=8)))
fig.add_hline(y=round(afrr_up_mean), line_dash='dash', line_color=CHART_COLORS['accent'],
              annotation_text=f'aFRR Up: {afrr_up_mean:.0f} MW')
fig.add_hline(y=round(total_bal), line_dash='dot', line_color=CHART_COLORS['primary'],
              annotation_text=f'Total Balancing: {total_bal:.0f} MW')
fig.add_vrect(x0=2024.8, x1=2026.2, fillcolor='rgba(39,174,96,0.07)', line_width=0)
fig.add_vrect(x0=2026.8, x1=2028.2, fillcolor='rgba(243,156,18,0.07)', line_width=0)
fig.add_vrect(x0=2028.8, x1=2030.2, fillcolor='rgba(192,57,43,0.07)', line_width=0)
fig.add_annotation(x=2025.5, y=100, text='Golden Window', showarrow=False, font=dict(size=10, color='green'))
fig.add_annotation(x=2027.5, y=100, text='Compression', showarrow=False, font=dict(size=10, color='orange'))
fig.add_annotation(x=2029.5, y=100, text='Saturated', showarrow=False, font=dict(size=10, color='red'))
fig.update_layout(title='BESS Build-Out Scenarios vs Balancing Market Size',
                  yaxis_title='MW', yaxis_range=[0, 1900], **layout_defaults)
charts['scenarios'] = fig_to_base64(fig)

# --- 15. Revenue vs Saturation ---
print("  15. Revenue vs Saturation")
bess_revenue = [0, 0, 120, 115, 110, 100, 85, 70, 55, 48, 42]
bess_sat_pct = [0, 10, 10, 12, 25, 38, 57, 71, 81, 86, 90]
fig = make_subplots(specs=[[{"secondary_y": True}]])
fig.add_trace(go.Bar(x=years_cap, y=bess_revenue, name='Revenue (EUR/kW/yr)',
                     marker_color=[CHART_COLORS['green'] if r > 100 else (CHART_COLORS['solar'] if r > 80 else CHART_COLORS['red']) for r in bess_revenue]),
              secondary_y=False)
fig.add_trace(go.Scatter(x=years_cap, y=bess_sat_pct, name='BESS/Peak Demand (%)',
                         mode='lines+markers', line=dict(color=CHART_COLORS['red'], width=3),
                         marker=dict(size=8)),
              secondary_y=True)
fig.update_layout(title='Revenue vs Saturation', **layout_defaults)
fig.update_yaxes(title_text='EUR/kW/yr', secondary_y=False)
fig.update_yaxes(title_text='BESS/Peak %', secondary_y=True)
charts['saturation'] = fig_to_base64(fig)

print(f"  Rendered {len(charts)} charts as SVG")


# ============================================================
# Build optimized HTML
# ============================================================
print("\nBuilding optimized HTML...")

# Read the current HTML for table data (we keep tables, replace chart divs)
with open(f"{OUT_DIR}/Lithuania_BESS_Market_Report.html", 'r', encoding='utf-8') as f:
    old_html = f.read()

# Extract table sections from old HTML (between specific markers)
# We'll rebuild from scratch for a clean optimized version

# Revenue tables data
rev_markets_all = ['DA Arbitrage', 'aFRR', 'FCR', 'mFRR', 'Imbalance', 'Multi-Market Combined']

def rev_table_rows(rev_dict, year_label):
    rows = ""
    for mkt in rev_markets_all:
        key_map = {'DA Arbitrage': 'DA Arbitrage', 'aFRR': 'aFRR', 'FCR': 'FCR',
                   'mFRR': 'mFRR', 'Imbalance': 'Imbalance', 'Multi-Market Combined': 'Combined'}
        k = key_map.get(mkt, mkt)
        vals = [rev_dict.get((k, d), 0) for d in durations]
        is_total = mkt == 'Multi-Market Combined'
        bold = 'font-weight:bold;' if is_total else ''
        bg = 'background:#D6DCE4;' if is_total else ''
        rows += f"<tr style='{bg}'>"
        rows += f"<td style='text-align:left;{bold}'>{mkt}</td>"
        for v in vals:
            rows += f"<td style='text-align:right;{bold}'>{v:,.0f}</td>"
        for v in vals:
            rows += f"<td style='text-align:right;{bold}'>{v/1000:.1f}</td>"
        rows += "</tr>"
    return rows

# Projection table rows
bess_installed = {2025: 454, 2026: 700, 2027: 1200, 2028: 1500, 2029: 1800, 2030: 2000}
proj_rows = ""
for year in range(2025, 2031):
    vals = [proj.get((year, d), 0) for d in durations]
    comp = COMPRESSION.get(year, 0.28)
    inst = bess_installed.get(year, '')
    proj_rows += f"<tr><td>{year}</td>"
    for v in vals: proj_rows += f"<td style='text-align:right'>{v:,.0f}</td>"
    for v in vals: proj_rows += f"<td style='text-align:right'>{v/1000:.1f}</td>"
    proj_rows += f"<td style='text-align:right'>{comp:.0%}</td>"
    proj_rows += f"<td style='text-align:right'>{inst}</td></tr>"

# DA Annual stats
da_stats_rows = ""
for year in range(2021, 2027):
    mask = da.index.year == year
    if mask.sum() == 0: continue
    v = da.loc[mask, 'price']
    daily = v.groupby(v.index.date)
    spread = daily.max() - daily.min()
    nh = int((v < 0).sum())
    da_stats_rows += f"""<tr><td><strong>{year}</strong></td>
        <td>{v.mean():.1f}</td><td>{v.min():.1f}</td><td>{v.max():.1f}</td>
        <td>{v.std():.1f}</td><td>{nh}</td></tr>"""

# Load stats
load_stats_rows = ""
for year in sorted(load_annual.keys()):
    la = load_annual[year]
    load_stats_rows += f"""<tr><td><strong>{year}</strong></td>
        <td>{la['avg']/1000:.2f}</td><td>{la['min']/1000:.2f}</td><td>{la['max']/1000:.2f}</td></tr>"""

# aFRR stats
afrr_stats_rows = ""
for year in [2024, 2025, 2026]:
    mask = afrr.index.year == year
    if mask.sum() == 0: continue
    yr = afrr.loc[mask]
    afrr_stats_rows += f"""<tr><td><strong>{year}</strong></td>
        <td>{yr['Up Prices'].mean():.2f}</td><td>{yr['Up Prices'].median():.2f}</td><td>{yr['Up Prices'].max():.0f}</td>
        <td>{yr['Down Prices'].mean():.2f}</td><td>{yr['Down Prices'].median():.2f}</td>
        <td>{yr['Up Quantity'].mean():.0f}</td></tr>"""

# Known projects
known_projects = [
    ('UAB Vėjo Galia', 53.6, 107.3, 'Kaišiadorys', 'Operational', 2025),
    ('European Energy', 25, 65, 'Anykščiai', 'Complete', 2026),
    ('Litgrid (TSO)', 200, 200, 'Various', 'Operational', 2024),
    ('Ignitis Group', 130, 260, 'TBD', 'Development', 2026),
    ('E Energija Group', 100, 200, 'TBD', 'Construction', 2026),
    ('Fluence / Litgrid', 50, 100, 'TBD', 'Integration', 2025),
]
projects_rows = ""
for dev, mw, mwh, loc, status, yr in known_projects:
    dur_h = mwh / mw if mw > 0 else 0
    sc = {'Operational': '#27ae60', 'Complete': '#2980b9', 'Construction': '#f39c12',
          'Development': '#e67e22', 'Integration': '#8e44ad'}.get(status, '#666')
    projects_rows += f"""<tr><td style="text-align:left">{dev}</td><td>{mw}</td><td>{mwh}</td>
        <td>{dur_h:.1f}</td><td>{loc}</td>
        <td><span style="background:{sc};color:white;padding:2px 8px;border-radius:10px;font-size:0.85em">{status}</span></td>
        <td>{yr}</td></tr>"""

# Conversion table
conv_rows = ""
for pct, new_mw, total, timeline in [(20, 340, 794, '2029-30'), (30, 510, 964, '2028-29'),
        (40, 680, 1134, '2028'), (60, 1020, 1474, '2027-28'), (80, 1360, 1814, '2027'), (100, 1700, 2154, '2028-30')]:
    sat = total / afrr_up_mean * 100
    sc = 'color:#c0392b;font-weight:bold' if sat > 200 else ('color:#e67e22' if sat > 100 else 'color:#27ae60')
    conv_rows += f"<tr><td>{pct}%</td><td>{new_mw:,} MW</td><td>{total:,} MW</td><td style='{sc}'>{sat:.0f}%</td><td>{timeline}</td></tr>"

# LinkedIn insights
linkedin_insights = [
    ('Balancing Services OÜ', 'Feb 5, 2025', 'Baltic balancing capacity launched. FCR hit EUR 145/MW/h. mFRR DOWN: EUR 20-30/MW/h. Lithuania joined for mFRR DOWN 440 MW.'),
    ('Fusebox Energy', 'Mar 2025', 'Baltic frequency reserves: EUR 9,976 to -EUR 4,473/MWh price swings. Single Latvian bidder caused distortions.'),
    ('Energy Lead', 'Aug 6, 2025', 'DA prices hit -EUR 3.06/MWh across Baltics. Ancillary services: -EUR 11,999/MWh in LV/LT.'),
    ('Zada (D. Zaitsev)', 'Jun 2025', '"A battery doing simple DA trading will be earning in excess of EUR 800/MW." Lithuania approved 4 GWh. "Won\'t last forever."'),
    ('Elektrum Eesti', 'Feb 2025', 'Estlink 2 failure: 65-71% Baltic price surge. Estonia covered only 64% of consumption locally.'),
    ('Litgrid', 'Jun 2025', 'First commercial BESS: 53.6 MW / 107.3 MWh. Total Lithuania: 453.9 MW / 461.5 MWh installed.'),
    ('Estonian Impact', 'Jan 2026', 'EUR 3.70/MWh reserve capacity fee + EUR 2/month "quiet sleep tax" for all consumers post-BRELL.'),
    ('VPPA/Spread', '2025', 'Baltic DA daily spread EUR 177. Analysis of 20M ENTSO-E rows: volatility is "fuel for battery profits."'),
    ('Grid Disconnect', 'Feb 7-9, 2025', 'BRELL disconnection. EUR 1.2B total cost (75% EU funded). TSOs deployed BESS for frequency management.'),
    ('Michael Dim', 'Sep 2025', '1.7 GW / 4 GWh pipeline. 50+ applications, EUR 840M+, 14.7% state subsidy. Max 300 MWh per facility.'),
    ('European Energy', 'Jan 2026', '25 MW / 65 MWh BESS completed in Anykščiai. Portfolio deployment approach across Baltics.'),
]
li_html = ""
for source, date, text in linkedin_insights:
    li_html += f"""<div style="background:#f8f9fa;border-left:4px solid #1F4E79;padding:12px 16px;margin:8px 0;border-radius:4px;">
        <strong>{source}</strong> <span style="color:#666;font-size:0.85em">({date})</span>
        <p style="margin:6px 0 0;color:#333;font-size:0.92em">{text}</p></div>"""

now = datetime.now().strftime('%B %d, %Y')

optimized_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Lithuania BESS Market Report — {now}</title>
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:Calibri,Arial,sans-serif;background:#f5f7fa;color:#333;line-height:1.6}}
.container{{max-width:1200px;margin:0 auto;padding:20px}}
header{{background:linear-gradient(135deg,#1F4E79 0%,#2E75B6 100%);color:white;padding:40px 30px;text-align:center;border-radius:12px;margin-bottom:30px}}
header h1{{font-size:2.2em;margin-bottom:8px}}
header p{{font-size:1.05em;opacity:0.9}}
h2{{color:#1F4E79;border-bottom:3px solid #1F4E79;padding-bottom:10px;margin:40px 0 20px}}
h3{{color:#2E75B6;margin:30px 0 15px}}
.card{{background:white;border-radius:8px;padding:24px;margin:16px 0;box-shadow:0 2px 8px rgba(0,0,0,0.08)}}
.chart-img{{text-align:center;margin:15px 0}}
.chart-img img{{max-width:100%;height:auto;border-radius:4px}}
table{{border-collapse:collapse;width:100%;font-size:0.9em;margin:10px 0}}
th{{background:#1F4E79;color:white;padding:10px;text-align:center}}
td{{padding:8px 12px;border:1px solid #ddd;text-align:center}}
tbody tr:nth-child(even){{background:#f2f7fb}}
tbody tr:hover{{background:#e2efda}}
.metric-grid{{display:grid;grid-template-columns:repeat(4,1fr);gap:20px;text-align:center}}
.metric-box{{background:linear-gradient(135deg,#2c3e50,#3498db);color:white;padding:20px;border-radius:10px}}
.metric-box .val{{font-size:2em;font-weight:bold}}
.metric-box .label{{color:#bdc3c7;font-size:0.9em}}
.timing-grid{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:16px;margin:20px 0}}
.timing-card{{border-radius:8px;padding:20px}}
.timing-card ul{{font-size:0.9em;padding-left:18px;margin:0}}
.takeaway-grid{{display:grid;grid-template-columns:1fr 1fr;gap:20px}}
.takeaway-box{{padding:15px}}
.takeaway-box h4{{color:#BDD7EE;margin:0 0 10px}}
.takeaway-box p{{margin:0;font-size:0.92em}}
details summary{{cursor:pointer;color:#2E75B6;font-weight:bold;margin-top:20px}}
details summary:hover{{text-decoration:underline}}
details>div{{padding:12px;background:#f8f9fa;border-radius:4px;margin-top:8px;font-size:0.88em;color:#555}}
footer{{text-align:center;padding:30px;color:#666;font-size:0.9em;margin-top:40px;border-top:2px solid #eee}}
@media(max-width:768px){{.metric-grid,.timing-grid,.takeaway-grid{{grid-template-columns:1fr}}.chart-img img{{width:100%}}}}
@media print{{body{{background:white}}.card{{box-shadow:none;border:1px solid #ddd}}header{{background:#1F4E79!important}}}}
</style>
</head>
<body>
<div class="container">

<header>
    <h1>Lithuania BESS Market Analysis</h1>
    <p>Battery Energy Storage Investment Analysis | Data: ENTSO-E Transparency Platform</p>
    <p style="font-size:0.85em;margin-top:8px;opacity:0.7">Generated {now} | Covers 2021-2026 with projections to 2030</p>
</header>

<!-- ===== SECTION 1: DAY-AHEAD PRICES ===== -->
<h2>1. Day-Ahead Electricity Prices</h2>

<div class="card">
    <h3>Monthly Average by Year</h3>
    <div class="chart-img">{img_tag(charts['da_monthly'], 'DA Monthly Prices')}</div>
</div>

<div class="card">
    <h3>Annual Statistics (EUR/MWh)</h3>
    <table>
        <thead><tr><th>Year</th><th>Mean</th><th>Min</th><th>Max</th><th>Std Dev</th><th>Neg Hours</th></tr></thead>
        <tbody>{da_stats_rows}</tbody>
    </table>
</div>

<div class="card">
    <h3>Daily Price Spread — Arbitrage Potential</h3>
    <div class="chart-img">{img_tag(charts['spread'], 'Daily Price Spread')}</div>
</div>

<div class="card">
    <h3>Hourly Price Profile</h3>
    <div class="chart-img">{img_tag(charts['hourly'], 'Hourly Profile')}</div>
</div>

<div class="card">
    <h3>Negative Price Hours</h3>
    <div class="chart-img">{img_tag(charts['neg_hours'], 'Negative Hours')}</div>
</div>

<!-- ===== SECTION 2: BESS REVENUE ===== -->
<h2>2. BESS Revenue Analysis by Duration &amp; Market</h2>

<div class="card">
    <h3>Revenue by Market Segment — 2025 Actual</h3>
    <div class="chart-img">{img_tag(charts['rev_2025'], 'Revenue 2025')}</div>
</div>

<div class="card">
    <h3>Detailed Revenue 2025 (EUR/MW/year)</h3>
    <table>
        <thead><tr><th style="text-align:left">Revenue Stream</th><th>1h<br>EUR/MW/yr</th><th>2h<br>EUR/MW/yr</th><th>4h<br>EUR/MW/yr</th><th>1h<br>EUR/kW/yr</th><th>2h<br>EUR/kW/yr</th><th>4h<br>EUR/kW/yr</th></tr></thead>
        <tbody>{rev_table_rows(rev_2025, '2025')}</tbody>
    </table>
</div>

<div class="card">
    <h3>Revenue 2024 — Pre-Desynchronization (EUR/MW/year)</h3>
    <table>
        <thead><tr><th style="text-align:left">Revenue Stream</th><th>1h<br>EUR/MW/yr</th><th>2h<br>EUR/MW/yr</th><th>4h<br>EUR/MW/yr</th><th>1h<br>EUR/kW/yr</th><th>2h<br>EUR/kW/yr</th><th>4h<br>EUR/kW/yr</th></tr></thead>
        <tbody>{rev_table_rows(rev_2024, '2024')}</tbody>
    </table>
</div>

<div class="card">
    <h3>Revenue Projection 2025-2030</h3>
    <div class="chart-img">{img_tag(charts['rev_projection'], 'Revenue Projection')}</div>
    <table style="margin-top:15px">
        <thead><tr><th>Year</th><th>1h EUR/MW/yr</th><th>2h EUR/MW/yr</th><th>4h EUR/MW/yr</th><th>1h EUR/kW/yr</th><th>2h EUR/kW/yr</th><th>4h EUR/kW/yr</th><th>Compression</th><th>BESS MW</th></tr></thead>
        <tbody>{proj_rows}</tbody>
    </table>
</div>

<div class="card" style="background:linear-gradient(135deg,#1F4E79 0%,#2E75B6 100%);color:white;">
    <h3 style="color:white;border:none;">Key Takeaways for BESS Investors</h3>
    <div class="takeaway-grid">
        <div class="takeaway-box"><h4>Extraordinary 2025 Revenues</h4><p>Post-BRELL scarcity: aFRR Up EUR 29.56/MW mean per 15-min (vs EUR 0.87 regulated in 2024). Baltic FCR EUR 145/MW/h on launch. Spikes to EUR 9,976/MWh.</p></div>
        <div class="takeaway-box"><h4>Revenue Compression Ahead</h4><p>454 MW installed → 2,000+ MW by 2030. "Opportunity won't last forever." Early movers capture 3-5x returns vs late entrants.</p></div>
        <div class="takeaway-box"><h4>Duration Matters</h4><p>4h BESS earns ~3.3x DA arbitrage of 1h. But aFRR/FCR capacity premium only ~1.4x. Optimal depends on market mix and timing.</p></div>
        <div class="takeaway-box"><h4>Structural Drivers Persist</h4><p>Baltic "energy island." Estlink 2 failures → 65-71% surges. 1.7 GW renewables added 2025. Reserve fees EUR 3.70/MWh create permanent pool.</p></div>
    </div>
</div>

<!-- ===== SECTION 3: PIPELINE & SATURATION ===== -->
<h2>3. Pipeline &amp; Market Saturation</h2>

<div class="metric-grid">
    <div class="metric-box"><div class="val">1.7 GW</div><div class="label">Pipeline Power</div></div>
    <div class="metric-box"><div class="val">4.0 GWh</div><div class="label">Pipeline Energy</div></div>
    <div class="metric-box"><div class="val">454 MW</div><div class="label">Installed Today</div></div>
    <div class="metric-box"><div class="val">50+</div><div class="label">Applications</div></div>
</div>
<div class="metric-grid" style="margin-top:12px">
    <div class="metric-box" style="background:linear-gradient(135deg,#34495e,#7f8c8d)"><div class="val">EUR 840M+</div><div class="label">Total Investment</div></div>
    <div class="metric-box" style="background:linear-gradient(135deg,#34495e,#7f8c8d)"><div class="val">14.7%</div><div class="label">Avg State Subsidy</div></div>
    <div class="metric-box" style="background:linear-gradient(135deg,#34495e,#7f8c8d)"><div class="val">~2.4h</div><div class="label">Avg Duration</div></div>
    <div class="metric-box" style="background:linear-gradient(135deg,#34495e,#7f8c8d)"><div class="val">300 MWh</div><div class="label">Max per Facility</div></div>
</div>

<div class="card">
    <h3>Pipeline vs Addressable Market</h3>
    <div class="chart-img">{img_tag(charts['pipeline'], 'Pipeline vs Market')}</div>
    <table style="margin-top:15px">
        <thead><tr><th style="text-align:left">Market Segment</th><th>2025 Actual (MW)</th><th>Pipeline Multiple</th><th>Assessment</th></tr></thead>
        <tbody>
            <tr><td style="text-align:left">aFRR Up</td><td>{afrr_up_mean:.0f}</td><td><strong>{1700/afrr_up_mean:.1f}x</strong></td><td style="color:#e67e22">Oversupplied if fully built</td></tr>
            <tr><td style="text-align:left">aFRR Down</td><td>{afrr_down_mean:.0f}</td><td>{1700/afrr_down_mean:.1f}x</td><td style="color:#e67e22">Oversupplied</td></tr>
            <tr><td style="text-align:left">mFRR</td><td>{mfrr_up_mean+mfrr_down_mean:.0f}</td><td>{1700/(mfrr_up_mean+mfrr_down_mean):.1f}x</td><td style="color:#c0392b">Heavily oversupplied</td></tr>
            <tr><td style="text-align:left">FCR (est.)</td><td>~40</td><td>42x</td><td style="color:#c0392b">Small market</td></tr>
            <tr style="font-weight:bold;background:#eee"><td style="text-align:left">Total Balancing</td><td>{total_bal:.0f}</td><td>{1700/total_bal:.1f}x</td><td style="color:#e67e22">Cannot absorb full pipeline</td></tr>
            <tr><td style="text-align:left">Peak Load</td><td>2,100</td><td>0.8x</td><td style="color:#27ae60">DA arbitrage structurally viable</td></tr>
        </tbody>
    </table>
</div>

<div class="card">
    <h3>Build-Out Scenarios</h3>
    <div class="chart-img">{img_tag(charts['scenarios'], 'Build-Out Scenarios')}</div>
</div>

<div class="card">
    <h3>Pipeline Conversion Probability</h3>
    <table>
        <thead><tr><th>Conversion</th><th>New Capacity</th><th>Total Installed</th><th>% of aFRR Up</th><th>Timeline</th></tr></thead>
        <tbody>{conv_rows}</tbody>
    </table>
</div>

<div class="timing-grid">
    <div class="timing-card" style="background:#e8f5e9;border:2px solid #27ae60">
        <h4 style="color:#27ae60;margin:0 0 10px">2025-2026: Golden Window</h4>
        <ul><li>454 MW vs 876 MW demand</li><li>Post-BRELL scarcity pricing</li><li>aFRR Up EUR 29.56/MW mean</li><li>Sub-1-year payback possible</li></ul>
    </div>
    <div class="timing-card" style="background:#fff3e0;border:2px solid #f39c12">
        <h4 style="color:#f39c12;margin:0 0 10px">2027-2028: Compression</h4>
        <ul><li>~950-1,200 MW installed (Base)</li><li>aFRR approaching saturation</li><li>Revenue 45-65% of 2025</li><li>2-3 year payback</li></ul>
    </div>
    <div class="timing-card" style="background:#fce4ec;border:2px solid #c0392b">
        <h4 style="color:#c0392b;margin:0 0 10px">2029-2030: Saturated</h4>
        <ul><li>1,400-1,700 MW installed</li><li>Balancing fully saturated</li><li>DA arbitrage still viable</li><li>Revenue 28-35% of 2025</li></ul>
    </div>
</div>

<div class="card">
    <h3>Known BESS Projects</h3>
    <table>
        <thead><tr><th style="text-align:left">Developer</th><th>MW</th><th>MWh</th><th>Duration</th><th>Location</th><th>Status</th><th>Year</th></tr></thead>
        <tbody>{projects_rows}</tbody>
    </table>
</div>

<!-- ===== SECTION 4: BALANCING MARKETS ===== -->
<h2>4. Balancing Market Data (aFRR / mFRR)</h2>

<div class="card">
    <h3>aFRR Contracted Reserve Prices — Monthly</h3>
    <div class="chart-img">{img_tag(charts['afrr_monthly'], 'aFRR Monthly')}</div>
    <table style="margin-top:15px">
        <thead><tr><th>Year</th><th>Up Mean</th><th>Up Median</th><th>Up Max</th><th>Down Mean</th><th>Down Median</th><th>Up Qty (MW)</th></tr></thead>
        <tbody>{afrr_stats_rows}</tbody>
    </table>
</div>

<div class="card">
    <h3>Imbalance Prices (EUR/MWh)</h3>
    <div class="chart-img">{img_tag(charts['imbalance'], 'Imbalance Prices')}</div>
    <p style="font-size:0.88em;color:#666;margin-top:8px">Note: Imbalance data available through Sep 2024 only. Post-PICASSO settlement uses 15-min ISP.</p>
</div>

<!-- ===== SECTION 5: LOAD & GENERATION ===== -->
<h2>5. Electricity Load &amp; Generation</h2>

<div class="card">
    <h3>Annual Load (GW)</h3>
    <div class="chart-img">{img_tag(charts['load'], 'Load')}</div>
    <table style="margin-top:10px">
        <thead><tr><th>Year</th><th>Avg (GW)</th><th>Min (GW)</th><th>Peak (GW)</th></tr></thead>
        <tbody>{load_stats_rows}</tbody>
    </table>
</div>

<div class="card">
    <h3>Generation by Source (TWh/yr)</h3>
    <div class="chart-img">{img_tag(charts['gen_stack'], 'Generation Mix')}</div>
</div>

<div class="card">
    <h3>Cross-Border Flows (Avg MW)</h3>
    <div class="chart-img">{img_tag(charts['flows'], 'Flows')}</div>
</div>

<!-- ===== SECTION 6: INSTALLED CAPACITY ===== -->
<h2>6. Installed Capacity &amp; Saturation</h2>

<div class="card">
    <h3>Generation Capacity (MW)</h3>
    <div class="chart-img">{img_tag(charts['capacity'], 'Installed Capacity')}</div>
</div>

<div class="card">
    <h3>Revenue vs BESS Saturation</h3>
    <div class="chart-img">{img_tag(charts['saturation'], 'Revenue vs Saturation')}</div>
</div>

<!-- ===== SECTION 7: MARKET INTELLIGENCE ===== -->
<h2>7. Baltic Market Intelligence (Industry Sources)</h2>
<div class="card">
    {li_html}
</div>

<div class="card">
    <h3>Structural Tailwinds</h3>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:20px">
        <div><strong style="color:#1F4E79">Growing Renewables</strong><p style="margin:4px 0;font-size:0.92em">Lithuania targeting 5.1 GW solar, doubled wind by 2030. 1.7 GW added 2025. More renewables = more volatility = more BESS opportunity.</p></div>
        <div><strong style="color:#1F4E79">Interconnection Risks</strong><p style="margin:4px 0;font-size:0.92em">Estlink 2 failure → 65-71% price surge. Baltics remain "energy island." Outages create extreme prices (EUR 9,976/MWh).</p></div>
        <div><strong style="color:#1F4E79">Baltic Market Integration</strong><p style="margin:4px 0;font-size:0.92em">Joint Baltic balancing procurement launched Feb 2025. Lithuanian BESS can serve LV/EE reserves. Combined aFRR need ~600-800 MW.</p></div>
        <div><strong style="color:#1F4E79">Reserve Capacity Fees</strong><p style="margin:4px 0;font-size:0.92em">EUR 3.70/MWh reserve fee (Estonia) creates permanent revenue pool for flexibility providers, independent of spot prices.</p></div>
    </div>
</div>

<!-- ===== METHODOLOGY ===== -->
<details>
    <summary>Methodology &amp; Assumptions</summary>
    <div>
        <ul>
            <li><strong>DA Arbitrage:</strong> Perfect-foresight upper bound (buy N cheapest, sell N most expensive hours/day) x 85% capture.</li>
            <li><strong>RT efficiency:</strong> 88% (Li-ion NMC/LFP).</li>
            <li><strong>aFRR:</strong> ENTSO-E contracted reserve prices (A47). Prices per MW per 15-min ISP. BESS alternates Up/Down (50/50) for SoC.</li>
            <li><strong>aFRR availability:</strong> 1h: 65%, 2h: 80%, 4h: 90%.</li>
            <li><strong>FCR:</strong> Estimated from Baltic market data (launched Feb 2025). EUR 30/MW/h avg (first-day: EUR 145/MW/h).</li>
            <li><strong>mFRR:</strong> ENTSO-E process type A51. Lower volumes, sporadic.</li>
            <li><strong>Imbalance:</strong> |Imbalance - DA| spread. Data through Sep 2024.</li>
            <li><strong>Multi-Market Combined:</strong> aFRR 40%, FCR 20%, DA 25%, mFRR 5%, Imbalance 10%.</li>
            <li><strong>Projections:</strong> Revenue compression as BESS grows 454 MW → 2,000+ MW. Balancing compresses faster than DA.</li>
            <li><strong>All figures gross revenue</strong> before opex (~EUR 5-8/kW/yr), degradation (2-3%/yr), financing.</li>
        </ul>
    </div>
</details>

<footer>
    <p>Lithuania BESS Market Analysis — Birdview Energy — {now}</p>
    <p>Data: ENTSO-E Transparency Platform | Industry: LinkedIn market intelligence</p>
    <p style="font-size:0.85em;margin-top:8px">Self-contained report. All charts embedded as SVG. No external dependencies.</p>
</footer>

</div>
</body>
</html>"""

# Write optimized HTML
output_path = f"{OUT_DIR}/Lithuania_BESS_Market_Report.html"
with open(output_path, 'w', encoding='utf-8') as f:
    f.write(optimized_html)

file_size = os.path.getsize(output_path)
print(f"\nOptimized report saved: {output_path}")
print(f"  File size: {file_size / 1024:.0f} KB ({file_size / 1024 / 1024:.1f} MB)")
print(f"  Charts: {len(charts)} SVG images embedded as base64")
print(f"  External dependencies: NONE (fully self-contained)")
print(f"  Works offline: YES")
print("Done!")
