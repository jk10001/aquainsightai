# filename: summary_statistics_table.py
import pandas as pd
import numpy as np

# Load data
file_name = "WWTP_microbial_loads_and_removal.xlsx"
xls = pd.ExcelFile(file_name)
df_data = pd.read_excel(xls, sheet_name="Data")
df_mdl = pd.read_excel(xls, sheet_name="Method detection limits")

# Normalize Plant codes: unify 'BiSP' -> 'BiSp'
df_data['Plant'] = df_data['Plant'].replace({'BiSP': 'BiSp'})

# Map IndicatorName to common microorganism names and Units from MDL sheet
mdl_units = dict(zip(df_mdl['Organism'].str.lower(), df_mdl['Units']))

def map_indicator_to_mdl_organism(ind_name):
    ind_name_lower = ind_name.lower()
    if 'somacoliphage' in ind_name_lower:
        return 'somatic coliphage'
    if 'malespeccoliphage' in ind_name_lower or 'male specific coliphage' in ind_name_lower:
        return 'male-specific coliphage'
    if 'aerobicendospore' in ind_name_lower or 'aerobic endospore' in ind_name_lower:
        return 'aerobic endospores'
    if 'totalcoliform' in ind_name_lower or 'total coliform' in ind_name_lower:
        return 'total coliform'
    if 'fecalcoliform' in ind_name_lower or 'fecal coliform' in ind_name_lower:
        return 'fecal coliform'
    if 'e.coli' in ind_name_lower or 'e coli' in ind_name_lower:
        return 'e. coli'
    if 'crypto' in ind_name_lower:
        return 'cryptosporidium'
    if 'giardia' in ind_name_lower:
        return 'giardia'
    if 'adenovirus' in ind_name_lower:
        return 'adenovirus'
    return None

df_data['Microorganism'] = df_data['IndicatorName'].apply(map_indicator_to_mdl_organism)
df_data['Units'] = df_data['Microorganism'].map(mdl_units).fillna('')

# Define microorganisms needing UF Effluent sample (effluent sample choice rule)
uf_effluent_organisms = {'giardia', 'cryptosporidium', 'somatic coliphage', 'male-specific coliphage'}

# Replace BDL and zero with organism-specific MDL, exclude N/A/NaN samples
# Convert all non-detects to MDL numeric for stats

# Create MDL lookup (lowercase keys)
mdl_lookup = {k.lower(): v for k, v in zip(df_mdl['Organism'].str.lower(), df_mdl['MDL'])}

def replace_bdl_zero_with_mdl(row):
    val = row['Count']
    if pd.isna(val):
        return np.nan  # sample not taken
    val_str = str(val).strip().upper()
    if val_str == 'BDL':
        return mdl_lookup.get(row['Microorganism'], np.nan)
    try:
        fval = float(val)
        if fval == 0:
            return mdl_lookup.get(row['Microorganism'], np.nan)
        else:
            return fval
    except:
        return np.nan  # cannot convert, treat as missing

df_data['Count_num'] = df_data.apply(replace_bdl_zero_with_mdl, axis=1)

# Select samples by stream and microorganism according to rules:
# Influent always sample type 'Influent grab'
df_influent = df_data[df_data['SampleType'] == 'Influent grab']

# Effluent samples:
df_uf_effluent = df_data[(df_data['SampleType'] == 'UF Effluent')]
df_grab_effluent = df_data[(df_data['SampleType'] == 'Effluent grab')]

# Build output rows list
rows = []

# Get unique plants and microorganisms
plants = sorted(df_data['Plant'].dropna().unique())
microorganisms = sorted(df_data['Microorganism'].dropna().unique())

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

# Compile function to get data by plant, microorganism, and stream
def get_stream_data(plant, microorganism, stream):
    if stream == 'Influent':
        df_sel = df_influent[(df_influent['Plant'] == plant) & (df_influent['Microorganism'] == microorganism)]
    else:  # Effluent
        if microorganism in uf_effluent_organisms:
            df_sel = df_uf_effluent[(df_uf_effluent['Plant'] == plant) & (df_uf_effluent['Microorganism'] == microorganism)]
        else:
            df_sel = df_grab_effluent[(df_grab_effluent['Plant'] == plant) & (df_grab_effluent['Microorganism'] == microorganism)]
    return df_sel['Count_num']

# Function to get units by microorganism
units_map = df_data.drop_duplicates(subset=['Microorganism'])[['Microorganism', 'Units']].set_index('Microorganism')['Units'].to_dict()

# Calculate stats for each plant, microorganism, and stream
for plant in plants:
    for microorganism in microorganisms:
        units = units_map.get(microorganism, '')
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
    units = units_map.get(microorganism, '')
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

# Create dataframe and format columns
df_summary = pd.DataFrame(rows)

# Rename microorganisms with consistent capitalization for output
microorganism_name_map = {
    'adenovirus': 'Adenovirus',
    'giardia': 'Giardia',
    'cryptosporidium': 'Cryptosporidium',
    'somatic coliphage': 'Somatic coliphage',
    'male-specific coliphage': 'Male-specific coliphage',
    'total coliform': 'Total coliform',
    'fecal coliform': 'Fecal coliform',
    'e. coli': 'E. coli',
    'aerobic endospores': 'Aerobic endospores'
}

def map_microorganism_name(name):
    return microorganism_name_map.get(name.lower(), name)

df_summary['Microorganism'] = df_summary['Microorganism'].apply(map_microorganism_name)

# Make sure counts are integers (because they are counts)
df_summary['n_samples [-]'] = df_summary['n_samples [-]'].astype(int)

# Sort output for neatness by Plant, Microorganism, Stream
df_summary = df_summary.sort_values(['Plant', 'Microorganism', 'Stream']).reset_index(drop=True)

# Save CSV
output_filename = "summary_statistics_influent_effluent_by_plant_pooled.csv"
df_summary.to_csv(output_filename, index=False)

print(output_filename)
print("summary_statistics_table.py")