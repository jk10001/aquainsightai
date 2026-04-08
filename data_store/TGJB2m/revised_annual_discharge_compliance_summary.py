# filename: revised_annual_discharge_compliance_summary.py
import pandas as pd
import numpy as np

# Data for the table
years = ['2014', '2015', '2016', '2017', '2018']

# Original data from the user input
flow_rates = [337.405, 342.521, 358.068, 354.452, 343.107]
ammonia_annual_mean = [0.0907258, 0.0859339, 0.0938031, 1.32132, 0.127017]
ammonia_90th = [0.1315, 0.11, 0.13, 0.13, 0.18]
bod_90th = [3.0, 4.0, 4.0, 3.0, 3.0]

# Discharge limits
limits = {
    'Flow Rate - Annual Mean (ML/d)': 540,
    'Ammonia - Annual Mean (mg/L)': 0.5,
    'Ammonia - 90th Percentile (mg/L)': 2.0,
    'BOD - 90th Percentile (mg/L)': 10.0,
}

# Function to round to 3 significant figures
def round_sig(x, sig=3):
    if x == 0:
        return 0
    return round(x, sig - int(np.floor(np.log10(abs(x)))) - 1)

# Round the values as required
flow_rates_rounded = [round_sig(x, 3) for x in flow_rates]
ammonia_annual_mean_rounded = [round_sig(x, 3) for x in ammonia_annual_mean]
ammonia_90th_rounded = [round_sig(x, 3) for x in ammonia_90th]

# BOD values remain as integers
bod_90th_int = [int(round(x)) for x in bod_90th]

# Construct the dataframe with parameters as rows
data = {
    'Parameter': ['Flow Rate - Annual Mean (ML/d)',
                  'Ammonia - Annual Mean (mg/L)',
                  'Ammonia - 90th Percentile (mg/L)',
                  'BOD - 90th Percentile (mg/L)'],
    'Discharge limit': [limits['Flow Rate - Annual Mean (ML/d)'],
                        limits['Ammonia - Annual Mean (mg/L)'],
                        limits['Ammonia - 90th Percentile (mg/L)'],
                        limits['BOD - 90th Percentile (mg/L)']],
    '2014': [flow_rates_rounded[0],
             ammonia_annual_mean_rounded[0],
             ammonia_90th_rounded[0],
             bod_90th_int[0]],
    '2015': [flow_rates_rounded[1],
             ammonia_annual_mean_rounded[1],
             ammonia_90th_rounded[1],
             bod_90th_int[1]],
    '2016': [flow_rates_rounded[2],
             ammonia_annual_mean_rounded[2],
             ammonia_90th_rounded[2],
             bod_90th_int[2]],
    '2017': [flow_rates_rounded[3],
             ammonia_annual_mean_rounded[3],
             ammonia_90th_rounded[3],
             bod_90th_int[3]],
    '2018': [flow_rates_rounded[4],
             ammonia_annual_mean_rounded[4],
             ammonia_90th_rounded[4],
             bod_90th_int[4]],
}

df = pd.DataFrame(data)

# Ensure columns order
df = df[['Parameter', 'Discharge limit', '2014', '2015', '2016', '2017', '2018']]

# Save the final processed table to CSV
output_csv = 'revised_annual_discharge_compliance_summary.csv'
df.to_csv(output_csv, index=False)

print(output_csv)