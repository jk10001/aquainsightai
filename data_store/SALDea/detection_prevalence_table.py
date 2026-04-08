# filename: detection_prevalence_table.py
import pandas as pd
import numpy as np

# Load data
file_name = "WWTP_microbial_loads_and_removal.xlsx"
xls = pd.ExcelFile(file_name)
df_data = pd.read_excel(xls, sheet_name="Data")
df_mdl = pd.read_excel(xls, sheet_name="Method detection limits")

# Normalize Plant codes: unify 'BiSP' -> 'BiSp'
df_data['Plant'] = df_data['Plant'].replace({'BiSP':'BiSp'})

# Map IndicatorName to Units from MDL sheet if possible
# Create a cleaned/common microorganism name from IndicatorName
# Because IndicatorName includes unit strings in some cases, also try to clean for common names
# We'll join Units from Method detection limits on Organism to IndicatorName

# Prepare a mapping from Organism to Units from MDL sheet and from Organism to cleaned Microorganism common name
# The Organism names in MDL sheet are more standardized, so try to map IndicatorName to Organism by contains or exact match

# Use MDL mapping dictionary for units
mdl_units = dict(zip(df_mdl['Organism'].str.lower(), df_mdl['Units']))

# Define element to map IndicatorName to Organism key in MDL for units
# Some IndicatorName values to Organism keys inferred:
# SomaColiphage -> Soma Coliphage
# MaleSpecColiphage -> Male Specific Coliphage
# AerobicEndospore -> Aerobic endospores
# TotalColiform -> Total coliform
# FecalColiform -> Fecal coliform
# E.coli -> E. coli
# CryptoOocyst -> Crypto Oocyst
# GiardiaCyst -> Giardia Cyst
# Adenovirus (MPN/L) -> Adenovirus

def map_indicator_to_mdl_organism(ind_name):
    ind_name_lower = ind_name.lower()
    if 'somacoliphage' in ind_name_lower:
        return 'soma coliphage'
    if 'malespeccoliphage' in ind_name_lower or 'male specific coliphage' in ind_name_lower:
        return 'male specific coliphage'
    if 'aerobicendospore' in ind_name_lower or 'aerobic endospore' in ind_name_lower:
        return 'aerobic endospores'
    if 'totalcoliform' in ind_name_lower or 'total coliform' in ind_name_lower:
        return 'total coliform'
    if 'fecalcoliform' in ind_name_lower or 'fecal coliform' in ind_name_lower:
        return 'fecal coliform'
    if 'e.coli' in ind_name_lower or 'e coli' in ind_name_lower:
        return 'e. coli'
    if 'crypto' in ind_name_lower:
        return 'crypto oocyst'
    if 'giardia' in ind_name_lower:
        return 'giardia cyst'
    if 'adenovirus' in ind_name_lower:
        return 'adenovirus'
    return None

# Create columns for Organism and Units
df_data['Organism_key'] = df_data['IndicatorName'].apply(map_indicator_to_mdl_organism)
df_data['Units'] = df_data['Organism_key'].map(mdl_units)
# For any not matched units, assign empty string to avoid NaNs
df_data['Units'] = df_data['Units'].fillna('')

# Define the four microorganisms needing UF Effluent for effluent sampling:
# Giardia, Cryptosporidium, Somatic coliphage, Male-specific coliphage as per problem statement
# Our Organism_key: 'giardia cyst', 'crypto oocyst', 'soma coliphage', 'male specific coliphage'
uf_effluent_organisms = {'giardia cyst', 'crypto oocyst', 'soma coliphage', 'male specific coliphage'}

# For detecting presence (non-detect definition):
# BDL and numeric 0 are non-detect
# NaN means sample not taken => exclude from denominator

def is_detected(val):
    if pd.isna(val):
        return None  # sample not taken
    val_str = str(val).strip().upper()
    if val_str == 'BDL':
        return False
    try:
        # Check numeric zero
        if float(val) == 0:
            return False
    except:
        # Not numeric
        pass
    # Otherwise present and nonzero means detected
    return True

# Apply detection boolean or None if no sample
df_data['Detected'] = df_data['Count'].apply(is_detected)

# Filter samples for Influent always 'Influent grab'
df_influent = df_data[df_data['SampleType'] == 'Influent grab'].copy()

# For Effluent, decide by organism
# Create two separate filters for effluent:
# For uf_effluent_organisms use SampleType == 'UF Effluent'
# For others use SampleType == 'Effluent grab'

# We'll split df_data into two effluent groups:
effluent_uf = df_data[df_data['SampleType'] == 'UF Effluent']
effluent_grab = df_data[df_data['SampleType'] == 'Effluent grab']

# We'll group results by Plant + Organism

# Prepare a function to calculate prevalence summary for a dataframe filtered for those sample types
def calc_prevalence(df_filtered, group_cols):
    # Counts number samples (excluding samples where Detected is None i.e. no samples)
    summary = df_filtered.groupby(group_cols).agg(
        n_samples = ('Detected', lambda x: x.notna().sum()),
        n_detected = ('Detected', lambda x: (x==True).sum())
    ).reset_index()
    # Calculate detection prevalence percent
    summary['detection_prevalence_percent'] = 100 * summary['n_detected'] / summary['n_samples']
    return summary

# Calculate prevalence for influent
influent_prev = calc_prevalence(df_influent, ['Plant', 'Organism_key', 'Units'])

# Calculate prevalence for effluent groups separately
effluent_uf_prev = calc_prevalence(effluent_uf, ['Plant', 'Organism_key', 'Units'])
effluent_grab_prev = calc_prevalence(effluent_grab, ['Plant', 'Organism_key', 'Units'])

# Join effluent prevalence by organism type based on uf_effluent_organisms list
# Build combined effluent prevalence by Plant, Organism_key, Units

# We want to choose effluent results per organism:
# If Organism_key in uf_effluent_organisms -> use effluent_uf_prev else effluent_grab_prev

# To do this, merge influent_prev with both effluent prevalence and choose appropriately

# First create a superset of all organism keys and plants from both influent and effluent
all_plants = sorted(df_data['Plant'].dropna().unique())
all_organisms = sorted(df_data['Organism_key'].dropna().unique())
all_units = df_data.drop_duplicates(subset=['Organism_key'])[['Organism_key', 'Units']].set_index('Organism_key')['Units'].to_dict()

index_tuples = []
for plant in all_plants:
    for org in all_organisms:
        index_tuples.append((plant, org))

df_index = pd.DataFrame(index_tuples, columns=['Plant', 'Organism_key'])
df_index['Units'] = df_index['Organism_key'].map(all_units).fillna('')

# Merge influent prevalence to index
df_index = df_index.merge(influent_prev[['Plant','Organism_key','Units','n_samples','n_detected','detection_prevalence_percent']],
                          on=['Plant','Organism_key','Units'], how='left')
df_index = df_index.rename(columns={
    'n_samples': 'n_influent_samples [-]',
    'n_detected': 'n_influent_detected [-]',
    'detection_prevalence_percent': 'influent_detection_prevalence [%]'
})

# Merge effluent grab prevalence
df_index = df_index.merge(effluent_grab_prev[['Plant','Organism_key','Units','n_samples','n_detected','detection_prevalence_percent']],
                          on=['Plant','Organism_key','Units'], how='left',
                          suffixes=(None,'_grab'))
df_index = df_index.rename(columns={
    'n_samples': 'n_effluent_samples_grab',
    'n_detected': 'n_effluent_detected_grab',
    'detection_prevalence_percent': 'effluent_detection_prevalence_grab'
})

# Merge effluent UF prevalence
df_index = df_index.merge(effluent_uf_prev[['Plant','Organism_key','Units','n_samples','n_detected','detection_prevalence_percent']],
                          on=['Plant','Organism_key','Units'], how='left',
                          suffixes=(None,'_uf'))
df_index = df_index.rename(columns={
    'n_samples': 'n_effluent_samples_uf',
    'n_detected': 'n_effluent_detected_uf',
    'detection_prevalence_percent': 'effluent_detection_prevalence_uf'
})

# Now select effluent columns according to organism rule
def select_effluent_row(row):
    if row['Organism_key'] in uf_effluent_organisms:
        # Use UF Effluent data
        n_samp = row['n_effluent_samples_uf']
        n_det = row['n_effluent_detected_uf']
        perc = row['effluent_detection_prevalence_uf']
    else:
        # Use Effluent grab data
        n_samp = row['n_effluent_samples_grab']
        n_det = row['n_effluent_detected_grab']
        perc = row['effluent_detection_prevalence_grab']
    # If nan, replace with 0 for counts and np.nan for percent if no samples
    n_samp = 0 if pd.isna(n_samp) else int(n_samp)
    n_det = 0 if pd.isna(n_det) else int(n_det)
    perc = 0.0 if pd.isna(perc) else perc
    return pd.Series({
        'n_effluent_samples [-]': n_samp,
        'n_effluent_detected [-]': n_det,
        'effluent_detection_prevalence [%]': perc
    })

effluent_selected = df_index.apply(select_effluent_row, axis=1)

# Combine final dataframe
df_final = pd.concat([
    df_index[['Plant', 'Organism_key', 'Units', 'n_influent_samples [-]', 'n_influent_detected [-]', 'influent_detection_prevalence [%]']],
    effluent_selected
], axis=1)

# Replace NaNs for influent with zeros for sample counts and detections, and zeros detection percent
df_final['n_influent_samples [-]'] = df_final['n_influent_samples [-]'].fillna(0).astype(int)
df_final['n_influent_detected [-]'] = df_final['n_influent_detected [-]'].fillna(0).astype(int)
df_final['influent_detection_prevalence [%]'] = df_final['influent_detection_prevalence [%]'].fillna(0)

# Rename Organism_key to Microorganism for final output
df_final = df_final.rename(columns={'Organism_key': 'Microorganism'})

# Sort by Plant then Microorganism
df_final = df_final.sort_values(['Plant', 'Microorganism']).reset_index(drop=True)

# Add "All plants pooled" row for each Microorganism
def pool_all_plants(df, organism, units):
    df_sub = df[df['Microorganism'] == organism]
    # Aggregate influent
    n_influent_samples = df_sub['n_influent_samples [-]'].sum()
    n_influent_detected = df_sub['n_influent_detected [-]'].sum()
    influent_prev = (100 * n_influent_detected / n_influent_samples) if n_influent_samples > 0 else 0
    # Aggregate effluent
    n_effluent_samples = df_sub['n_effluent_samples [-]'].sum()
    n_effluent_detected = df_sub['n_effluent_detected [-]'].sum()
    effluent_prev = (100 * n_effluent_detected / n_effluent_samples) if n_effluent_samples > 0 else 0
    return {
        'Plant': 'All plants pooled',
        'Microorganism': organism,
        'Units': units,
        'n_influent_samples [-]': n_influent_samples,
        'n_influent_detected [-]': n_influent_detected,
        'influent_detection_prevalence [%]': influent_prev,
        'n_effluent_samples [-]': n_effluent_samples,
        'n_effluent_detected [-]': n_effluent_detected,
        'effluent_detection_prevalence [%]': effluent_prev
    }

pooled_rows = []
for organism in df_final['Microorganism'].unique():
    units = df_final.loc[df_final['Microorganism'] == organism, 'Units'].iloc[0]
    pooled_rows.append(pool_all_plants(df_final, organism, units))

df_pooled = pd.DataFrame(pooled_rows)

# Append pooled data to df_final
df_output = pd.concat([df_final, df_pooled], ignore_index=True)

# Save to CSV
base_filename = "detection_prevalence_influent_vs_effluent_by_plant_pooled"
csv_filename = base_filename + ".csv"
df_output.to_csv(csv_filename, index=False)

print(csv_filename)
print("detection_prevalence_table.py")