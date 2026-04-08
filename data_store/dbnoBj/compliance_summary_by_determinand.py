# filename: compliance_summary_by_determinand.py
import pandas as pd
import numpy as np

# Load data
file_name = "Anglian_Water_Domestic_Water_Quality.csv"
df = pd.read_csv(file_name, encoding="utf-8")

# Pre-processing
df['Sample_Date'] = pd.to_datetime(df['Sample_Date'], format='%d/%m/%Y %H:%M', errors='coerce')
df = df[df['Determinand'].notna() & df['Result'].notna()].copy()
# Operator value '<' means left-censored, but for compliance we use Result value as is
# Replace any Operator values other than '<' with NaN to avoid confusion (not needed here)
df['Operator'] = df['Operator'].replace({'<': '<'}).where(df['Operator'] == '<')

# Mapping dataset Determinand -> regulatory determinand, PCV value and units
# Include non-numeric PCVs as None for numeric fields but with labels
mapping = {
    # Microbiological / indicator
    'Coliform Bacteria (Indicator)': ('Coliform bacteria', 0, 'number/100ml', True),
    'E.Coli (faecal coliforms Confirmed)': ('Escherichia coli', 0, 'number/100ml', True),
    'Enterococci (Confirmed)': ('Enterococci', 0, 'number/100ml', True),
    'Clostridum Perfringens (Sulphite-reducing Clostridia) (Confirmed)': ('Clostridium perfringens', 0, 'Number/100ml', True),
    'Colour': ('Colour', 20, 'mg/l Pt/Co', True),
    'Turbidity': ('Turbidity', 1, 'NTU', True),
    'Odour': ('Odour', None, '', False),
    'Taste (Taste Quant)': ('Taste', None, '', False),
    'Hydrogen ion (pH) - Indicator (Hydrogen ion) (pH)': ('pH', None, '', False),

    # Inorganic / chemical
    'Nitrate (Total)': ('Nitrate', 50, 'mgNO3/l', True),
    "Nitrite - Consumer's Taps": ('Nitrite', 0.50, 'mgNO2/l', True),
    'Nitrite/Nitrate formula': ('Nitrate+Nitrite formula', 1, '', True),  # special combined assessment, treat as numeric compliance with PCV=1
    'Ammonium (Total)': ('Ammonium', 0.50, 'mgNH4/l', True),
    'Aluminium (Total)': ('Aluminium', 200, 'µg/l', True),
    'Iron (Total)': ('Iron', 200, 'µg/l', True),
    'Manganese (Total)': ('Manganese', 50, 'µg/l', True),
    'Conductivity (Electrical Conductivity)': ('Conductivity', 2500, 'µS/cm @ 20°C', True),
    'Sodium (Total)': ('Sodium', 200, 'mg/l', True),
    'Chloride': ('Chloride', 250, 'mg/l', True),
    'Sulphate': ('Sulphate', 250, 'mg/l', True),
    'Total Organic Carbon': ('TOC', None, '', False),
    'Boron': ('Boron', 1.0, 'mg/l', True),
    'Fluoride (Total)': ('Fluoride', 1.5, 'mg/l', True),
    'Cyanide (Total)': ('Cyanide', 50, 'µg/l', True),
    'Mercury (Total)': ('Mercury', 1.0, 'µg/l', True),
    'Nickel (Total)': ('Nickel', 20, 'µg/l', True),
    'Lead (10 - will apply 25.12.2013)': ('Lead', 10, 'µg/l', True),
    'Antimony': ('Antimony', 5.0, 'µg/l', True),
    'Selenium (Total)': ('Selenium', 10, 'µg/l', True),
    'Copper (Total)': ('Copper', 2.0, 'mg/l', True),
    'Chromium (Total)': ('Chromium', 50, 'µg/l', True),

    # Disinfection by-products
    'Bromate': ('Bromate', 10, 'µg/l', True),
    'Trihalomethanes (Total by Calculation)': ('THM (total)', 100, 'µg/l', True),

    # Radiological
    'Gross Alpha': ('gross alpha', 0.1, 'Bq/l', True),
    'Gross Beta': ('gross beta', 1, 'Bq/l', True),
    'Tritium': ('tritium', 100, 'Bq/l', True),

    # PAH / organics
    'Polycyclic Aromatic Hydrocarbons (Total by Calculation)': ('PAH', 0.10, 'µg/l', True),
    'Benzo[a]Pyrene (Total)': ('benzo[a]pyrene', 0.010, 'µg/l', True),

    # Pesticides
    'Pesticides Aldrin': ('Aldrin', 0.030, 'µg/l', True),
    'Pesticides Dieldrin (Total)': ('Dieldrin', 0.030, 'µg/l', True),
    'Pesticides Heptachlor (Total)': ('Heptachlor', 0.10, 'µg/l', True),
    'Pesticides Heptachlor Epoxide - Total (Trans, CIS) (Heptachlor Epoxide)': ('Heptachlor epoxide', 0.10, 'µg/l', True),
    'Pesticides (Total by Calculation)': ('Pesticides: total', 0.50, 'µg/l', True)
}
# For other pesticides matching pattern "Pesticides XXX"
# PCV = 0.10 µg/l, name = "Other pesticides"
def map_pesticide(determinand):
    # Check known pesticide keys above first
    if determinand in mapping:
        return mapping[determinand]
    if determinand.startswith('Pesticides '):
        return ('Other pesticides', 0.10, 'µg/l', True)
    return None

def get_mapping(determinand):
    res = mapping.get(determinand)
    if res is not None:
        return res
    pest_map = map_pesticide(determinand)
    if pest_map is not None:
        return pest_map
    # Not mapped
    return (None, None, None, False)

# Add regulatory info to df
df[['Reg_Determinand_Name', 'Reg_PCV_Value', 'Reg_PCV_Units', 'Is_Numeric_PCV']] = df['Determinand'].apply(
    lambda x: pd.Series(get_mapping(x))
)

# Special case: handling 'Nitrite/Nitrate formula' compliance calculation requires both nitrate and nitrite.
# So we will calculate compliance separately for that and for other determinands.

# For Nitrite/Nitrate formula, calculate formula = (Nitrate/50) + (Nitrite/3) per Sample_Id & Sample_Date,
# but dataset likely has separate rows for Nitrite/Nitrate formula, so we just use reported Result for that determinand.

# Prepare compliance summary list
summary_rows = []

# Aggregate over determinand groups
grouped = df.groupby('Determinand')

for determinand, group in grouped:
    reg_name = group['Reg_Determinand_Name'].iloc[0]
    pcv_value = group['Reg_PCV_Value'].iloc[0]
    pcv_units = group['Reg_PCV_Units'].iloc[0]
    is_numeric = group['Is_Numeric_PCV'].iloc[0]

    total_samples = len(group)

    if not is_numeric or reg_name is None:
        # Non-numeric PCVs or not assessed: blanks for numeric fields
        compliance_pass_fail = ''
        num_fail = ''
        mean_result = ''
        max_result = ''
    else:
        # For numeric PCVs: evaluate compliance
        # For micro params with PCV = 0, fail if result > 0, else fail if result > PCV
        if pcv_value == 0:
            num_fail = (group['Result'] > 0).sum()
        else:
            num_fail = (group['Result'] > pcv_value).sum()

        compliance_pass_fail = 'Pass' if num_fail == 0 else 'Fail'
        mean_result = group['Result'].mean()
        max_result = group['Result'].max()

    summary_rows.append({
        'Determinand': determinand,
        'Reg_Determinand_Name': reg_name if reg_name else '',
        'Reg_PCV_Value': pcv_value if pcv_value is not None else '',
        'Reg_PCV_Units': pcv_units if pcv_units else '',
        'Compliance_Pass_Fail': compliance_pass_fail,
        'Total_Samples': total_samples if compliance_pass_fail != '' else '',
        'Num_Fail': num_fail if compliance_pass_fail != '' else '',
        'Mean_Result': mean_result if compliance_pass_fail != '' else '',
        'Max_Result': max_result if compliance_pass_fail != '' else ''
    })

# Create summary dataframe
summary_df = pd.DataFrame(summary_rows)

# Save to CSV
output_file = "compliance_summary_by_determinand.csv"
summary_df.to_csv(output_file, index=False)

print(output_file)
print("compliance_summary_by_determinand.py")