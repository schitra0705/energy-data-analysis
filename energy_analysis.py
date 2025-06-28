import pandas as pd

# Load the Excel file
df = pd.read_excel("DailySummaryTest.xlsx")

# 2.1 Site with Highest DG Run Hours
max_dg_site = df.loc[df['DG1_RUNHRS_CONS'].idxmax()]
site_id = max_dg_site['Site ID']
dg_run_hours = round(max_dg_site['DG1_RUNHRS_CONS'], 2)
print(f"2.1 Site with Highest DG Run Hours:\nSite ID: {site_id}, DG Run Hours: {dg_run_hours}\n")

# 2.2 Top 5 Sites with Highest Grid Run Hours
top5_grid_sites = df.nlargest(5, 'EB_RUNHRS_CONS')[['Site ID', 'EB_RUNHRS_CONS']]
print("2.2 Top 5 Sites with Highest Grid Run Hours:")
print(top5_grid_sites.to_string(index=False))
print()

# 2.3 Total Grid Consumption
total_grid_kwh = df['EB_KWH_CONS'].sum()
print(f"2.3 Total Grid Consumption (KWH) for all 50 sites:\n{round(total_grid_kwh, 2)} kWh\n")

# 2.4 Summary of Sites Based on POWER MODEL (M1, M5, M7)
df['POWER_MODEL_TYPE'] = df['POWER MODEL'].str.extract(r'(M[157])')
model_summary = df['POWER_MODEL_TYPE'].value_counts()

print("2.4 Summary of Sites Based on POWER MODEL (M1, M5, M7):")
for model in ['M1', 'M5', 'M7']:
    print(f"{model}: {model_summary.get(model, 0)}")
