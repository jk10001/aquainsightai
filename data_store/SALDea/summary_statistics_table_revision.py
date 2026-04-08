# filename: summary_statistics_table_revision.py
import pandas as pd
import numpy as np

# Load data
file_name = "WWTP_microbial_loads_and_removal.xlsx"
xls = pd.ExcelFile(file_name)
df_data = pd.read_excel(xls, sheet_name="Data")
df_mdl = pd.read_excel(xls, sheet_name="Method detection limits")

# Normalize Plant codes: unify 'BiSP' -> 'BiSp'
df_data['Plant'] = df_data['Plant'].replace({'BiSP': 'BiSp'})

# Mapping from output Microorganism names to MDL sheet Organism names for consistent Units and MDL lookup
microorganism_to_mdl_organism = {
    'Adenovirus': 'Adenovirus',
    'Giardia': 'Giardia Cyst',
    'Cryptosporidium': 'Crypto Oocyst',
    'Somatic coliphage': 'Soma Coliphage',
    'Male-specific coliphage': 'Male Specific Coliphage',
    'Total coliform': 'Total coliform',
    'Fecal coliform': 'Fecal coliform',
    'E. coli': 'E. coli',
    'Aerobic endospores': 'Aerobic endospores',
}

# Original mapping from IndicatorName to output Microorganism names (lowercase keys mapped to desired output)
indicator_to_microorganism = {
    'somacoliphage': 'Somatic coliphage',
    'malespeccoliphage': 'Male-specific coliphage',
    'aerobicendospore': 'Aerobic endospores',
    'totalcoliform': 'Total coliform',
    'fecalcoliform': 'Fecal coliform',
    'e.coli': 'E. coli',
    'crypto': 'Cryptosporidium',
    'giardia': 'Giardia',
    'adenovirus': 'Adenovirus',
}

def map_indicator_to_microorganism(ind_name):
    ind_name_lower = ind_name.lower()
    for key in indicator_to_microorganism:
        if key in ind_name_lower:
            return indicator_to_microorganism[key]
    return None

df_data['Microorganism'] = df_data['IndicatorName'].apply(map_indicator_to_microorganism)

# Create a reverse dict from mdl data for MDL and Units keyed by Organism name
mdl_dict = {}
for _, row in df_mdl.iterrows():
    mdl_dict[row['Organism']] = {'MDL': row['MDL'], 'Units': row['Units']}

# Add columns for MDL and Units by mapping Microorganism to MDL Organism names
def get_mdl_and_units(microorganism):
    mdl_org = microorganism_to_mdl_organism.get(microorganism)
    if mdl_org and mdl_org in mdl_dict:
        return mdl_dict[mdl_org]['MDL'], mdl_dict[mdl_org]['Units']
    else:
        return np.nan, ''

df_data[['MDL', 'Units']] = df_data['Microorganism'].apply(
    lambda x: pd.Series(get_mdl_and_units(x))
)

# Define microorganisms needing UF Effluent sample (effluent sample choice rule)
uf_effluent_organisms = {'Giardia', 'Cryptosporidium', 'Somatic coliphage', 'Male-specific coliphage'}

# Replace BDL and numeric zero with organism-specific MDL, exclude N/A/NaN samples
def replace_bdl_zero_with_mdl(row):
    val = row['Count']
    if pd.isna(val):
        return np.nan  # sample not taken
    val_str = str(val).strip().upper()
    if val_str == 'BDL':
        return row['MDL']
    try:
        fval = float(val)
        if fval == 0:
            return row['MDL']
        else:
            return fval
    except:
        return np.nan  # cannot convert, treat as missing

df_data['Count_num'] = df_data.apply(replace_bdl_zero_with_mdl, axis=1)

# Select samples by stream and microorganism according to rules:
# Influent always sample type 'Influent grab'
df_influent = df_data[df_data['SampleType'] == 'Influent grab']

# Effluent samples:
df_uf_effluent = df_data[df_data['SampleType'] == 'UF Effluent']
df_grab_effluent = df_data[df_data['SampleType'] == 'Effluent grab']

# Prepare function to compute summary stats
def compute_summary_stats(series):
    ser = series.dropna()
    n = ser.count()
    if n == 0:
        return {'n_samples [-]': 0, 'median [original units]': np.nan, 'mean [original units]': np.nan,
                'p25 [original units]': np.nan, 'p75 [original units]': np.nan}
    return {'n_samples [-]': n,
            'median [original units]': ser.median(),
            'mean [original units]': ser.mean(),
            'p25 [original units]': ser.quantile(0.25),
            'p75 [original units]': ser.quantile(0.75)}

# Get data by plant, microorganism, stream
def get_stream_data(plant, microorganism, stream):
    if stream == 'Influent':
        df_sel = df_influent[(df_influent['Plant'] == plant) & (df_influent['Microorganism'] == microorganism)]
    else:  # Effluent
        if microorganism in uf_effluent_organisms:
            df_sel = df_uf_effluent[(df_uf_effluent['Plant'] == plant) & (df_uf_effluent['Microorganism'] == microorganism)]
        else:
            df_sel = df_grab_effluent[(df_grab_effluent['Plant'] == plant) & (df_grab_effluent['Microorganism'] == microorganism)]
    return df_sel['Count_num']

# Unique plants and microorganisms
plants = sorted(df_data['Plant'].dropna().unique())
microorganisms = sorted(df_data['Microorganism'].dropna().unique())

rows = []

# Calculate stats for each plant, microorganism, stream
for plant in plants:
    for microorganism in microorganisms:
        units = ''
        mdl_org_name = microorganism_to_mdl_organism.get(microorganism)
        if mdl_org_name and mdl_org_name in mdl_dict:
            units = mdl_dict[mdl_org_name]['Units']
        # Influent
        influent_vals = get_stream_data(plant, microorganism, 'Influent')
        influent_stats = compute_summary_stats(influent_vals)
        rows.append({
            'Plant': plant,
            'Microorganism': microorganism,
            'Units': units,
            'Stream': 'Influent',
            **influent_stats
        })
        # Effluent
        effluent_vals = get_stream_data(plant, microorganism, 'Effluent')
        effluent_stats = compute_summary_stats(effluent_vals)
        rows.append({
            'Plant': plant,
            'Microorganism': microorganism,
            'Units': units,
            'Stream': 'Effluent',
            **effluent_stats
        })

# Add pooled "All plants pooled" rows for each microorganism and stream
for microorganism in microorganisms:
    units = ''
    mdl_org_name = microorganism_to_mdl_organism.get(microorganism)
    if mdl_org_name and mdl_org_name in mdl_dict:
        units = mdl_dict[mdl_org_name]['Units']
    # Influent pooled
    influent_vals = pd.concat([get_stream_data(plant, microorganism, 'Influent') for plant in plants], ignore_index=True)
    influent_stats = compute_summary_stats(influent_vals)
    rows.append({
        'Plant': 'All plants pooled',
        'Microorganism': microorganism,
        'Units': units,
        'Stream': 'Influent',
        **influent_stats
    })
    # Effluent pooled
    effluent_vals = pd.concat([get_stream_data(plant, microorganism, 'Effluent') for plant in plants], ignore_index=True)
    effluent_stats = compute_summary_stats(effluent_vals)
    rows.append({
        'Plant': 'All plants pooled',
        'Microorganism': microorganism,
        'Units': units,
        'Stream': 'Effluent',
        **effluent_stats
    })

# Create dataframe and sort
df_summary = pd.DataFrame(rows)

# Sort for neatness
df_summary = df_summary.sort_values(['Plant', 'Microorganism', 'Stream']).reset_index(drop=True)

# Make sure counts are integers
df_summary['n_samples [-]'] = df_summary['n_samples [-]'].astype(int)

# Save the corrected CSV
output_filename = "summary_statistics_influent_effluent_by_plant_pooled.csv"
df_summary.to_csv(output_filename, index=False)

print(output_filename)
print("summary_statistics_table_revision.py")