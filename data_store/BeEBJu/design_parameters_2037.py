# filename: design_parameters_2037.py
import pandas as pd

# Filenames
pop_proj_csv = "population_projection_20yr.csv"
flows_proj_csv = "historical_projected_flows_20yr.csv"
avg_loads_proj_csv = "historical_projected_avg_loads_BOD_TSS_TP_20yr.csv"
max_month_loads_proj_csv = "historical_projected_max_month_loads_BOD_TSS_TP_20yr.csv"

# Target year
target_year = 2037

# Load projected flows and filter to target year
df_flows = pd.read_csv(flows_proj_csv)
df_flows_2037 = df_flows[df_flows["Year"] == target_year]

# Load projected average day loads and filter
df_avg_loads = pd.read_csv(avg_loads_proj_csv)
df_avg_loads_2037 = df_avg_loads[df_avg_loads["Year"] == target_year]

# Load projected max month loads and filter
df_max_month_loads = pd.read_csv(max_month_loads_proj_csv)
df_max_month_loads_2037 = df_max_month_loads[df_max_month_loads["Year"] == target_year]

# Extract required parameter values
avg_day_flow = df_flows_2037["Avg_flow_projected_m3_d"].values[0] if not df_flows_2037.empty else None
max_day_flow = df_flows_2037["Max_flow_projected_m3_d"].values[0] if not df_flows_2037.empty else None

avg_bod_load = df_avg_loads_2037["BOD_avg_load_kg_d_projected"].values[0] if not df_avg_loads_2037.empty else None
avg_tss_load = df_avg_loads_2037["TSS_avg_load_kg_d_projected"].values[0] if not df_avg_loads_2037.empty else None
avg_tp_load = df_avg_loads_2037["TP_avg_load_kg_d_projected"].values[0] if not df_avg_loads_2037.empty else None

max_month_bod_load = df_max_month_loads_2037["BOD_max_month_load_kg_d_projected"].values[0] if not df_max_month_loads_2037.empty else None
max_month_tss_load = df_max_month_loads_2037["TSS_max_month_load_kg_d_projected"].values[0] if not df_max_month_loads_2037.empty else None
max_month_tp_load = df_max_month_loads_2037["TP_max_month_load_kg_d_projected"].values[0] if not df_max_month_loads_2037.empty else None

# Round values according to instructions
def round_flow(val):
    return round(val / 10) * 10 if pd.notna(val) else None

def round_load(val):
    return round(val / 10) * 10 if pd.notna(val) else None

avg_day_flow_r = round_flow(avg_day_flow)
max_day_flow_r = round_flow(max_day_flow)

avg_bod_load_r = round_load(avg_bod_load)
avg_tss_load_r = round_load(avg_tss_load)
avg_tp_load_r = round_load(avg_tp_load)

max_month_bod_load_r = round_load(max_month_bod_load)
max_month_tss_load_r = round_load(max_month_tss_load)
max_month_tp_load_r = round_load(max_month_tp_load)

# Prepare output table DataFrame
rows = [
    ["Average day flow", avg_day_flow_r, "m³/d"],
    ["Maximum day flow", max_day_flow_r, "m³/d"],
    ["Average day BOD load", avg_bod_load_r, "kg/d"],
    ["Average day TSS load", avg_tss_load_r, "kg/d"],
    ["Average day Total P load", avg_tp_load_r, "kg/d"],
    ["Maximum month BOD load (30-day avg)", max_month_bod_load_r, "kg/d"],
    ["Maximum month TSS load (30-day avg)", max_month_tss_load_r, "kg/d"],
    ["Maximum month Total P load (30-day avg)", max_month_tp_load_r, "kg/d"],
]

df_output = pd.DataFrame(rows, columns=["Parameter", "Value", "Units"])

# Save to CSV
output_csv = "design_parameters_2037.csv"
df_output.to_csv(output_csv, index=False)

# Print filenames
print("design_parameters_2037.py")
print(output_csv)