import json
import math
import os

base_dir = r"d:\Perkuliahan\Kelass\SKB"
cfg_path = os.path.join(base_dir, "04_INTERNAL_TOOLS_DAN_DATA", "03_data_dan_aset", "config_bakmi_mimu.json")
if not os.path.exists(cfg_path):
    cfg_path = os.path.join(base_dir, "config_bakmi_mimu.json")

with open(cfg_path, "r", encoding="utf-8") as f:
    cfg = json.load(f)

# 1. Capex
capex_total = sum(item["qty"] * item["unit_price"] for item in cfg["capex_items"])
sewa = cfg["operational_costs"]["rent_per_year"]
amdal = cfg["operational_costs"]["amdal_initial"]
survey = cfg["operational_costs"]["survey_initial"]
promosi = cfg["operational_costs"]["promosi_initial"]
initial_outlay = capex_total + sewa + amdal + survey + promosi
depresiasi_per_year = capex_total * cfg["financial_assumptions"]["depreciation_rate"]

print(f"Capex Total: Rp {capex_total:,.0f}")
print(f"Initial Outlay: Rp {initial_outlay:,.0f}")
print(f"Depresiasi per Tahun (20%): Rp {depresiasi_per_year:,.0f}")

# 2. Inflow per year
# Days = 25 * 12 = 300 days
cif_years = [0] * 5
products = cfg["products"]
for yr in range(5):
    yr_total = 0
    for prod in products:
        sales_qty = prod["daily_sales"][yr]
        # Price
        if yr == 0:
            price = prod["price_year_1"]
        else:
            # Escalation: Match Excel ROUNDUP(..., -3)
            prev_price = prod["price_year_1"]
            for y in range(1, yr + 1):
                rate = prod["price_increase_rates"][y]
                prev_price = math.ceil((prev_price * (1 + rate)) / 1000) * 1000
            price = prev_price
        monthly_rev = sales_qty * price * 25
        annual_rev = monthly_rev * 12
        yr_total += annual_rev
    cif_years[yr] = yr_total

print("\nCASH INFLOW 5 TAHUN:")
for i, c in enumerate(cif_years, start=1):
    print(f"  Tahun {i}: Rp {c:,.0f}")

# 3. Outflow per year
# Gaji, bahan baku, listrik, gas, air, galon
base_gaji = sum(emp["count"] * emp["monthly_salary"] * 12 for emp in cfg["operational_costs"]["employees"])
base_bb = cfg["operational_costs"]["raw_materials_monthly_base"] * 12
base_listrik = cfg["operational_costs"]["electricity_monthly_base"] * 12
base_air = cfg["operational_costs"]["water_monthly_base"] * 12
base_gas = cfg["operational_costs"]["gas_monthly_base"] * 12
base_galon = 400000 * 12

cof_years = [0] * 5
for yr in range(5):
    # Gaji
    rate_g = (0.10 + yr * 0.02) if yr > 0 else 0
    gaji_yr = base_gaji if yr == 0 else cof_years[yr-1]["gaji"] * (1 + rate_g)
    
    # BB
    rate_bb = (0.12 + yr * 0.03) if yr > 0 else 0
    bb_yr = base_bb if yr == 0 else cof_years[yr-1]["bb"] * (1 + rate_bb)
    
    # Listrik
    rate_l = (0.10 + yr * 0.02) if yr > 0 else 0
    l_yr = base_listrik if yr == 0 else cof_years[yr-1]["listrik"] * (1 + rate_l)
    
    # Air
    rate_a = (0.08 + yr * 0.02) if yr > 0 else 0
    a_yr = base_air if yr == 0 else cof_years[yr-1]["air"] * (1 + rate_a)
    
    # Gas
    gas_yr = base_gas
    
    # Galon
    galon_yr = base_galon + (yr * 5 * 20000)

    tot_cof = gaji_yr + bb_yr + l_yr + a_yr + gas_yr + galon_yr
    cof_years[yr] = {
        "total": tot_cof,
        "gaji": gaji_yr,
        "bb": bb_yr,
        "listrik": l_yr,
        "air": a_yr,
        "gas": gas_yr,
        "galon": galon_yr
    }

print("\nCASH OUTFLOW 5 TAHUN:")
for i, co in enumerate(cof_years, start=1):
    print(f"  Tahun {i}: Rp {co['total']:,.0f}")

# 4. NCF, EAT, Proceed
ncf_years = [cif_years[i] - cof_years[i]["total"] for i in range(5)]
proceed_years = [ncf_years[i] - depresiasi_per_year for i in range(5)]

print("\nNET CASH FLOW (NCF) 5 TAHUN:")
for i, n in enumerate(ncf_years, start=1):
    print(f"  Tahun {i}: Rp {n:,.0f}")

print("\nPROCEED (NCF - Depresiasi) 5 TAHUN:")
for i, p in enumerate(proceed_years, start=1):
    print(f"  Tahun {i}: Rp {p:,.0f}")

# 5. Discounted Cash Flow (Rate = 20%)
discount_rate = 0.20
pv_years = [proceed_years[i] / ((1 + discount_rate) ** (i + 1)) for i in range(5)]
total_pv = sum(pv_years)
npv = total_pv - initial_outlay

print(f"\nTOTAL PV: Rp {total_pv:,.0f}")
print(f"NPV: Rp {npv:,.0f} (Status: {'LAYAK' if npv > 0 else 'TIDAK LAYAK'})")

# 6. IRR
# Approximation via numpy_financial or bisection
def calc_npv(rate):
    return sum(proceed_years[i] / ((1 + rate) ** (i + 1)) for i in range(5)) - initial_outlay

low = 0.0
high = 10.0
for _ in range(100):
    mid = (low + high) / 2
    if calc_npv(mid) > 0:
        low = mid
    else:
        high = mid
irr = mid
print(f"IRR: {irr * 100:.2f}% (Status: {'LAYAK' if irr > discount_rate else 'TIDAK LAYAK'})")

# 7. Payback Period
# In year 1, proceed vs initial outlay
pp_months = (initial_outlay / (proceed_years[0] / 12))
print(f"Payback Period: {pp_months:.2f} Bulan ({pp_months/12:.2f} Tahun) (Status: {'LAYAK' if pp_months < 36 else 'TIDAK LAYAK'})")

# 8. Profitability Index
pi = sum(proceed_years) / initial_outlay
print(f"Profitability Index: {pi:.2f} (Status: {'LAYAK' if pi > 1.20 else 'TIDAK LAYAK'})")

# Save results to json for docx and pptx generator
metrics = {
    "capex_total": capex_total,
    "initial_outlay": initial_outlay,
    "depresiasi_per_year": depresiasi_per_year,
    "cif_years": cif_years,
    "cof_years": [c["total"] for c in cof_years],
    "ncf_years": ncf_years,
    "proceed_years": proceed_years,
    "pv_years": pv_years,
    "total_pv": total_pv,
    "npv": npv,
    "irr": irr,
    "pp_months": pp_months,
    "pp_years": pp_months / 12,
    "pi": pi
}

met_path = os.path.join(base_dir, "04_INTERNAL_TOOLS_DAN_DATA", "03_data_dan_aset", "metrics_bakmi_mimu.json")
with open(met_path, "w", encoding="utf-8") as f:
    json.dump(metrics, f, indent=2)

# Also write to root for backward compatibility
root_met_path = os.path.join(base_dir, "metrics_bakmi_mimu.json")
try:
    with open(root_met_path, "w", encoding="utf-8") as f:
        json.dump(metrics, f, indent=2)
except Exception:
    pass

print(f"\n[SUCCESS] Metrics computed and saved to: {met_path}")
