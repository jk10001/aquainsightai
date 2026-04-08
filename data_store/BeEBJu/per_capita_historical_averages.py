# filename: per_capita_historical_averages.py
import pandas as pd

# Load CSVs
flow_csv = "annual_per_capita_flow_avg_max.csv"
avg_load_csv = "annual_per_capita_avg_loads_BOD_TSS_TP.csv"
max_month_load_csv = "max_month_per_capita_loads_BOD_TSS_TP.csv"

df_flow = pd.read_csv(flow_csv)
df_avg_load = pd.read_csv(avg_load_csv)
df_max_month_load = pd.read_csv(max_month_load_csv)

# Extract historical averages ("Average" row)
flow_avg_row = df_flow[df_flow["Year_or_Group"] == "Average"].iloc[0]
avg_load_avg_row = df_avg_load[df_avg_load["Year_or_Group"] == "Average"].iloc[0]
max_month_load_avg_row = df_max_month_load[df_max_month_load["Year_or_Group"] == "Average"].iloc[0]

# Prepare data rows
rows = [
    # Flow per capita
    ["Average annual daily flow per capita", round(flow_avg_row["Average_annual_daily_flow_L_per_p_d"], 1), "L/person/day"],
    ["Maximum day flow per capita", round(flow_avg_row["Maximum_day_flow_L_per_p_d"], 1), "L/person/day"],
    
    # Loads per capita - average day
    ["Average day BOD load per capita", round(avg_load_avg_row["BOD_avg_daily_load_g_p_d"], 1), "g/person/day"],
    ["Average day TSS load per capita", round(avg_load_avg_row["TSS_avg_daily_load_g_p_d"], 1), "g/person/day"],
    ["Average day Total P load per capita", round(avg_load_avg_row["TP_avg_daily_load_g_p_d"], 1), "g/person/day"],
    
    # Loads per capita - maximum month
    ["Maximum month BOD load per capita (30-day avg)", round(max_month_load_avg_row["BOD_max_month_daily_load_g_p_d"], 1), "g/person/day"],
    ["Maximum month TSS load per capita (30-day avg)", round(max_month_load_avg_row["TSS_max_month_daily_load_g_p_d"], 1), "g/person/day"],
    ["Maximum month Total P load per capita (30-day avg)", round(max_month_load_avg_row["TP_max_month_daily_load_g_p_d"], 1), "g/person/day"]
]

# Create DataFrame for output
df_output = pd.DataFrame(rows, columns=["Parameter", "Value", "Units"])

# Save CSV
output_csv = "per_capita_historical_averages.csv"
df_output.to_csv(output_csv, index=False)

# Print filenames to terminal
print(output_csv)
print("per_capita_historical_averages.py")