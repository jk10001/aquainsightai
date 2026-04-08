# filename: annual_discharge_compliance.py
import pandas as pd

file_name = "MWC_ETP_Data.xlsx"

# Load sheets needed
xls = pd.ExcelFile(file_name)
df_influent_flow = pd.read_excel(xls, sheet_name='ETP Influent Flow')
df_effluent_quality = pd.read_excel(xls, sheet_name='ETP Effluent Quality')
df_discharge_limits = pd.read_excel(xls, sheet_name='Discharge Limits')

# Extract relevant discharge limits into variables for easier use and clarity
limit_flow_annual_mean = df_discharge_limits.loc[df_discharge_limits['Parameter'] == 'Flow Rate - annual mean', 'Discharge Limit'].values[0]
limit_ammonia_annual_mean = df_discharge_limits.loc[df_discharge_limits['Parameter'] == 'Ammonia - annual mean', 'Discharge Limit'].values[0]
limit_ammonia_90th = df_discharge_limits.loc[df_discharge_limits['Parameter'] == 'Ammonia - 90th percentile', 'Discharge Limit'].values[0]
limit_bod_90th = df_discharge_limits.loc[df_discharge_limits['Parameter'] == 'BOD - 90th percentile', 'Discharge Limit'].values[0]

# Prepare Influent Flow Data for annual mean flow rate as proxy for discharge flow
df_influent_flow['Year'] = df_influent_flow['Date'].dt.year
annual_flow_mean = df_influent_flow.groupby('Year')['Influent Flow (ML/d)'].mean().reset_index()
annual_flow_mean.rename(columns={'Influent Flow (ML/d)': 'Flow Rate - Annual Mean (ML/d)'}, inplace=True)

# Prepare Effluent Quality data
df_effluent_quality['Year'] = df_effluent_quality['Date'].dt.year

# Calculate ammonia annual mean per year
ammonia_annual_mean = df_effluent_quality.groupby('Year')['Ammonia (mg/L)'].mean().reset_index()
ammonia_annual_mean.rename(columns={'Ammonia (mg/L)': 'Ammonia - Annual Mean (mg/L)'}, inplace=True)

# Calculate ammonia 90th percentile per year
ammonia_90th = df_effluent_quality.groupby('Year')['Ammonia (mg/L)'].quantile(0.9).reset_index()
ammonia_90th.rename(columns={'Ammonia (mg/L)': 'Ammonia - 90th Percentile (mg/L)'}, inplace=True)

# Calculate BOD 90th percentile per year
bod_90th = df_effluent_quality.groupby('Year')['BOD (mg/L)'].quantile(0.9).reset_index()
bod_90th.rename(columns={'BOD (mg/L)': 'BOD - 90th Percentile (mg/L)'}, inplace=True)

# Merge all measures into a single dataframe on Year
df_annual_summary = annual_flow_mean.merge(ammonia_annual_mean, on='Year', how='left')
df_annual_summary = df_annual_summary.merge(ammonia_90th, on='Year', how='left')
df_annual_summary = df_annual_summary.merge(bod_90th, on='Year', how='left')

# Add the corresponding limit columns for each parameter
df_annual_summary['Flow Rate Limit (ML/d)'] = limit_flow_annual_mean
df_annual_summary['Ammonia Annual Mean Limit (mg/L)'] = limit_ammonia_annual_mean
df_annual_summary['Ammonia 90th Percentile Limit (mg/L)'] = limit_ammonia_90th
df_annual_summary['BOD 90th Percentile Limit (mg/L)'] = limit_bod_90th

# Reorder columns as requested
df_annual_summary = df_annual_summary[['Year',
                                       'Flow Rate - Annual Mean (ML/d)', 'Flow Rate Limit (ML/d)',
                                       'Ammonia - Annual Mean (mg/L)', 'Ammonia Annual Mean Limit (mg/L)',
                                       'Ammonia - 90th Percentile (mg/L)', 'Ammonia 90th Percentile Limit (mg/L)',
                                       'BOD - 90th Percentile (mg/L)', 'BOD 90th Percentile Limit (mg/L)']]

# Save to csv
output_filename = 'annual_discharge_compliance_summary_2014_2018.csv'
df_annual_summary.to_csv(output_filename, index=False)

print(output_filename)