# Lithuania BESS Market Analysis

Comprehensive Battery Energy Storage System (BESS) investment analysis for the Lithuanian electricity market, covering revenue modeling, market dynamics, and saturation assessment.

## Key Findings

| Metric | Value |
|--------|-------|
| **2025 Multi-Market Revenue (4h BESS)** | ~578,000 EUR/MW/year |
| **2030 Projected (with saturation)** | ~198,000 EUR/MW/year |
| **BESS Installed (2025)** | 454 MW / 462 MWh |
| **BESS Pipeline** | 1.7 GW / 4.0 GWh pipeline (50+ applications) |
| **aFRR Requirement** | ~410 MW (2025 actual) |
| **DA Daily Price Spread** | EUR 177 mean (2025) |

## Context

Lithuania disconnected from the Russian BRELL grid on February 8-9, 2025 and synchronized with Continental Europe. This created:
- Extreme scarcity in balancing reserves (aFRR Up prices jumped from EUR 0.87 to EUR 29.56 mean per 15-min period)
- FCR market launch with EUR 145/MW/h clearing on day one
- Frequency reserve price spikes from EUR 9,976 to -EUR 4,473/MWh
- 65-71% electricity price increases during Estlink 2 cable failure

These conditions make Lithuania one of Europe's most attractive BESS markets in 2025, but revenue compression is expected as 4 GW of pipeline capacity enters the market.

## Outputs

### Interactive HTML Report
**`Lithuania_BESS_Market_Report.html`** — Open in any browser. Contains:
- Day-ahead price analysis with hourly profiles and seasonal patterns
- Generation mix breakdown (wind, solar, gas, hydro)
- Cross-border flow analysis (Sweden NordBalt, Poland LitPol, Latvia)
- Extended balancing market data (aFRR/mFRR post-PICASSO)
- **BESS revenue analysis by duration (1h/2h/4h) across 6 market segments**
- Forward revenue projections 2025-2030 with saturation effects
- Baltic market intelligence from industry sources

### Excel Workbook
**`BirdEnergySystemInstalled_Lithuania.xlsx`** — Contains sheets:
1. **Installed Capacity** — Wind, Solar PV, BESS, Fossil (2021-2030 historical + forecast)
2. **ENTSO-E Real Data** — Live API data analysis with monthly/annual statistics
3. **BESS Revenue Analysis** — Revenue by duration and market segment with projections
4. **Balancing Data (API)** — aFRR/mFRR contracted reserve prices and activated energy
5. **Day-Ahead Prices** — Historical and forecast EUR/MWh
6. **Electricity Load** — TWh/yr, GW avg/min/max
7. **Balancing & Ancillary** — aFRR, mFRR, FCR market overview
8. **BESS Saturation Analysis** — Pipeline vs market size assessment
9. **Market Overview** — Key developments and regulatory changes

## Data Sources

### ENTSO-E Transparency Platform (via API)
20 CSV files in `data/` covering 2021-2026:

| File | Description | Rows |
|------|-------------|------|
| `da_prices_LT.csv` | Day-ahead hourly prices | ~54,000 |
| `afrr_reserve_prices_LT.csv` | aFRR contracted reserve prices (15-min, A47) | ~56,000 |
| `mfrr_reserve_prices_LT.csv` | mFRR contracted reserve prices (15-min, A51) | ~56,000 |
| `imbalance_prices_LT.csv` | Imbalance settlement prices | ~33,000 |
| `actual_load_LT.csv` | System load | ~77,000 |
| `generation_by_type_LT.csv` | Generation by fuel type | ~77,000 |
| `flow_{X}_to_{Y}.csv` | Cross-border flows (SE_4, PL, LV) | ~6 files |
| `installed_capacity_LT_{year}.csv` | Installed generation capacity | 6 files |
| `activated_balancing_energy_prices_LT.csv` | Activated energy prices | ~2,300 |

### Industry Sources (LinkedIn)
Market intelligence from Balancing Services OU, Fusebox Energy, Litgrid, Zada, Elektrum Eesti, European Energy, and others — integrated into the HTML report and Excel.

## Scripts

| Script | Purpose |
|--------|---------|
| `fetch_entsoe_data.py` | Retrieve core market data from ENTSO-E API |
| `fetch_balancing_extended.py` | Retrieve aFRR/mFRR reserve prices (post-Sep 2024 gap) |
| `create_lithuania_bess_analysis.py` | Generate Excel workbook with research data |
| `build_report.py` | Build interactive HTML report from ENTSO-E data |
| `update_report_with_balancing.py` | Add extended balancing data section to report |
| `add_revenue_section.py` | Add BESS revenue analysis by duration and market |
| `add_pipeline_section.py` | Add pipeline & saturation analysis with build-out scenarios |

## Revenue Model

Revenue computed from actual ENTSO-E data for 1h, 2h, and 4h BESS across:

- **DA Arbitrage** — Buy cheapest N hours, sell most expensive N hours (85% capture rate)
- **aFRR** — Contracted reserve capacity prices (ENTSO-E A47 endpoint)
- **FCR** — Estimated from Baltic market data (launched Feb 2025)
- **mFRR** — Contracted reserve capacity prices (ENTSO-E A51 endpoint)
- **Imbalance** — DA vs imbalance price spread trading
- **Multi-Market Combined** — Optimized time allocation across all markets

### Key Assumptions
- Round-trip efficiency: 88%
- aFRR availability: 65% (1h) / 80% (2h) / 90% (4h)
- Multi-market allocation: aFRR 40%, FCR 20%, DA 25%, mFRR 5%, Imbalance 10%
- Revenue compression: 454 MW installed (2025) growing to 2,000+ MW (2030)
- All figures are gross revenue before opex, degradation, and financing

## Setup

### Requirements
```
pip install pandas numpy openpyxl xlsxwriter entsoe-py python-dotenv requests plotly
```

### API Key
Place your ENTSO-E API key in `~/.env`:
```
ENTSOE_API_KEY=your_key_here
```
Register at https://transparency.entsoe.eu/ and request API access via email.

### Running
```bash
# 1. Fetch data from ENTSO-E (takes ~5 minutes)
python fetch_entsoe_data.py
python fetch_balancing_extended.py

# 2. Generate Excel workbook
python create_lithuania_bess_analysis.py

# 3. Build HTML report and update Excel with live data
python build_report.py

# 4. Add balancing market section
python update_report_with_balancing.py

# 5. Add revenue analysis section
python add_revenue_section.py

# 6. Open report
open Lithuania_BESS_Market_Report.html
```

## Known Data Gaps

- **Imbalance prices post-Sep 2024**: ENTSO-E stopped publishing after Baltic 15-min ISP transition. Bridged with aFRR/mFRR reserve price endpoints.
- **Intraday prices**: Not available on ENTSO-E for Lithuania (only via Nord Pool proprietary API).
- **FCR prices**: No ENTSO-E data; estimated from industry sources (Baltic FCR market launched Feb 2025).
- **Installed capacity**: ENTSO-E returns NaN for Lithuania; uses web-researched data.

## License

Internal analysis. Data sourced from ENTSO-E Transparency Platform (public) and industry publications.
