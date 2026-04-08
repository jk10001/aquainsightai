# filename: lrv_summary_table.py
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

# Add MDL column by mapping Microorganism to MDL Organism names
def get_mdl(microorganism):
    mdl_org = microorganism_to_mdl_organism.get(microorganism)
    if mdl_org and mdl_org in mdl_dict:
        return mdl_dict[mdl_org]['MDL']
    else:
        return np.nan

df_data['MDL'] = df_data['Microorganism'].apply(get_mdl)

# Define microorganisms needing UF Effluent sample (effluent sample choice rule)
uf_effluent_organisms = {'Giardia', 'Cryptosporidium', 'Somatic coliphage', 'Male-specific coliphage'}

# Replace BDL and numeric zero with organism-specific MDL for calculations
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

# Filter data by sample types for influent and effluent by microorganism
df_influent = df_data[df_data['SampleType'] == 'Influent grab'].copy()
df_uf_effluent = df_data[df_data['SampleType'] == 'UF Effluent'].copy()
df_grab_effluent = df_data[df_data['SampleType'] == 'Effluent grab'].copy()

# We need to pair influent and effluent samples by Plant + Event + Microorganism
# Select effluent sample type per organism rules
def select_effluent_samples(microorganism):
    if microorganism in uf_effluent_organisms:
        return df_uf_effluent
    else:
        return df_grab_effluent

# Prepare list for results
rows = []

# Get unique Plants and Microorganisms
plants = sorted(df_data['Plant'].dropna().unique())
microorganisms = sorted(df_data['Microorganism'].dropna().unique())

# For each plant and microorganism, we will:
# - get influent samples (Plant, Event, Microorganism, Count_num)
# - get effluent samples (Plant, Event, Microorganism, Count_num)
# - merge on Plant + Event + Microorganism to pair samples
# - calculate LRV where influent_count > MDL (skip if influent_count == MDL)
# - summarize LRV (count, median, mean, 25th, 75th percentile)

for plant in plants:
    for microorganism in microorganisms:
        mdl_value = get_mdl(microorganism)
        if pd.isna(mdl_value):
            # skip if no MDL available
            continue

        # Prepare influent subset
        infl_sub = df_influent[(df_influent['Plant'] == plant) & (df_influent['Microorganism'] == microorganism)][['Event','Count_num']]
        infl_sub = infl_sub.rename(columns={'Count_num': 'Influent_count'})

        # Select appropriate effluent samples per microorganism
        df_effluent_sel = select_effluent_samples(microorganism)
        eff_sub = df_effluent_sel[(df_effluent_sel['Plant'] == plant) & (df_effluent_sel['Microorganism'] == microorganism)][['Event','Count_num']]
        eff_sub = eff_sub.rename(columns={'Count_num': 'Effluent_count'})

        # Merge on Event to pair samples for same event in same plant & microorganism
        paired = pd.merge(infl_sub, eff_sub, on='Event', how='inner')

        # Drop pairs with nan effluent or influent counts
        paired = paired.dropna(subset=['Influent_count', 'Effluent_count'])

        # Skip pairs where influent count == MDL (no reliable removal calculation if influent is considered at detection limit)
        paired = paired[paired['Influent_count'] > mdl_value]

        # Calculate LRV = log10(Influent / Effluent)
        # Avoid division by zero or invalid
        paired['LRV'] = np.log10(paired['Influent_count'] / paired['Effluent_count'])

        # Remove infinite or negative infinite LRV values if any (could arise if Effluent_count is zero, but zeros replaced with MDL so unlikely)
        paired = paired[np.isfinite(paired['LRV'])]

        n_pairs = len(paired)
        if n_pairs == 0:
            median_lrv = np.nan
            mean_lrv = np.nan
            p25_lrv = np.nan
            p75_lrv = np.nan
        else:
            median_lrv = paired['LRV'].median()
            mean_lrv = paired['LRV'].mean()
            p25_lrv = paired['LRV'].quantile(0.25)
            p75_lrv = paired['LRV'].quantile(0.75)

        rows.append({
            'Plant': plant,
            'Microorganism': microorganism,
            'n_pairs [-]': n_pairs,
            'median_LRV [log10]': median_lrv,
            'mean_LRV [log10]': mean_lrv,
            'p25_LRV [log10]': p25_lrv,
            'p75_LRV [log10]': p75_lrv
        })

# Add pooled 'All plants pooled' rows
for microorganism in microorganisms:
    mdl_value = get_mdl(microorganism)
    if pd.isna(mdl_value):
        continue

    infl_list = []
    eff_list = []
    for plant in plants:
        infl_sub = df_influent[(df_influent['Plant'] == plant) & (df_influent['Microorganism'] == microorganism)][['Event','Count_num']]
        infl_sub = infl_sub.rename(columns={'Count_num': 'Influent_count'})
        df_effluent_sel = select_effluent_samples(microorganism)
        eff_sub = df_effluent_sel[(df_effluent_sel['Plant'] == plant) & (df_effluent_sel['Microorganism'] == microorganism)][['Event','Count_num']]
        eff_sub = eff_sub.rename(columns={'Count_num': 'Effluent_count'})

        paired = pd.merge(infl_sub, eff_sub, on='Event', how='inner')
        paired = paired.dropna(subset=['Influent_count','Effluent_count'])
        paired = paired[paired['Influent_count'] > mdl_value]

        infl_list.append(paired['Influent_count'])
        eff_list.append(paired['Effluent_count'])

    if infl_list:
        infl_concat = pd.concat(infl_list, ignore_index=True)
        eff_concat = pd.concat(eff_list, ignore_index=True)
        paired_all = pd.DataFrame({'Influent_count': infl_concat, 'Effluent_count': eff_concat})
        # Calculate LRV
        paired_all['LRV'] = np.log10(paired_all['Influent_count'] / paired_all['Effluent_count'])
        paired_all = paired_all[np.isfinite(paired_all['LRV'])]
        n_pairs = len(paired_all)

        if n_pairs == 0:
            median_lrv = np.nan
            mean_lrv = np.nan
            p25_lrv = np.nan
            p75_lrv = np.nan
        else:
            median_lrv = paired_all['LRV'].median()
            mean_lrv = paired_all['LRV'].mean()
            p25_lrv = paired_all['LRV'].quantile(0.25)
            p75_lrv = paired_all['LRV'].quantile(0.75)

        rows.append({
            'Plant': 'All plants pooled',
            'Microorganism': microorganism,
            'n_pairs [-]': n_pairs,
            'median_LRV [log10]': median_lrv,
            'mean_LRV [log10]': mean_lrv,
            'p25_LRV [log10]': p25_lrv,
            'p75_LRV [log10]': p75_lrv
        })

# Create output DataFrame and round LRV statistics to 1 decimal place
df_lrv_summary = pd.DataFrame(rows)
lrvs = ['median_LRV [log10]', 'mean_LRV [log10]', 'p25_LRV [log10]', 'p75_LRV [log10]']
for col in lrvs:
    df_lrv_summary[col] = df_lrv_summary[col].round(1)

# Sort output by Plant then Microorganism
df_lrv_summary = df_lrv_summary.sort_values(['Plant', 'Microorganism']).reset_index(drop=True)

# Save CSV
output_filename = "log_removal_value_summary_by_plant_pooled.csv"
df_lrv_summary.to_csv(output_filename, index=False)

print(output_filename)
print("lrv_summary_table.py")