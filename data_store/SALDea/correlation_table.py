# filename: correlation_table.py
import pandas as pd
import numpy as np
from scipy.stats import spearmanr, pearsonr

file_name = "WWTP_microbial_loads_and_removal.xlsx"
xls = pd.ExcelFile(file_name)
df_data = pd.read_excel(xls, sheet_name="Data")
df_mdl = pd.read_excel(xls, sheet_name="Method detection limits")

# Normalize plant codes
df_data['Plant'] = df_data['Plant'].replace({'BiSP': 'BiSp'})

# Define pathogens and indicators
pathogens = ['Adenovirus', 'Giardia', 'Cryptosporidium']
indicators = [
    'Total coliform', 'Fecal coliform', 'E. coli',
    'Somatic coliphage', 'Male-specific coliphage', 'Aerobic endospores'
]

# Map IndicatorName to final Microorganism names for pathogens and indicators combined
indicator_map = {
    'totalcoliform': 'Total coliform',
    'fecalcoliform': 'Fecal coliform',
    'e.coli': 'E. coli',
    'somacoliphage': 'Somatic coliphage',
    'malespeccoliphage': 'Male-specific coliphage',
    'aerobicendospore': 'Aerobic endospores',
    'giardia': 'Giardia',
    'crypto': 'Cryptosporidium',
    'adenovirus': 'Adenovirus',
}

def map_indicator_to_microorganism(ind_name):
    ind_name_lower = ind_name.lower()
    for key, val in indicator_map.items():
        if key in ind_name_lower:
            return val
    return None

df_data['Microorganism'] = df_data['IndicatorName'].apply(map_indicator_to_microorganism)
df_data = df_data[df_data['Microorganism'].notna()]

# MDL dictionary
mdl_dict = {row['Organism'].lower(): row['MDL'] for _, row in df_mdl.iterrows()}

# Mapping Microorganism to MDL organism name
microorganism_to_mdl_organism = {
    'Adenovirus': 'Adenovirus',
    'Giardia': 'Giardia Cyst',
    'Cryptosporidium': 'Crypto Oocyst',
    'Total coliform': 'Total coliform',
    'Fecal coliform': 'Fecal coliform',
    'E. coli': 'E. coli',
    'Somatic coliphage': 'Soma Coliphage',
    'Male-specific coliphage': 'Male Specific Coliphage',
    'Aerobic endospores': 'Aerobic endospores',
}

def get_mdl(microorganism):
    mdl_org = microorganism_to_mdl_organism.get(microorganism)
    if mdl_org:
        return mdl_dict.get(mdl_org.lower(), np.nan)
    else:
        return np.nan

df_data['MDL'] = df_data['Microorganism'].apply(get_mdl)

# Replace BDL and 0 with MDL for count_num
def replace_bdl_zero_with_mdl(row):
    val = row['Count']
    if pd.isna(val):
        return np.nan
    val_str = str(val).strip().upper()
    if val_str == 'BDL':
        return row['MDL']
    try:
        fval = float(val)
        if fval == 0:
            return row['MDL']
        return fval
    except:
        return np.nan

df_data['Count_num'] = df_data.apply(replace_bdl_zero_with_mdl, axis=1)

# Effluent selection rules
uf_effluent_organisms = {'Giardia', 'Cryptosporidium', 'Somatic coliphage', 'Male-specific coliphage'}

# Split data by sample type
df_influent = df_data[df_data['SampleType'] == 'Influent grab']
df_uf_effluent = df_data[df_data['SampleType'] == 'UF Effluent']
df_grab_effluent = df_data[df_data['SampleType'] == 'Effluent grab']

def get_effluent_df(microorganism):
    if microorganism in uf_effluent_organisms:
        return df_uf_effluent[df_uf_effluent['Microorganism'] == microorganism]
    else:
        return df_grab_effluent[df_grab_effluent['Microorganism'] == microorganism]

def get_lrv_pairs(microorganism):
    df_inf = df_influent[df_influent['Microorganism'] == microorganism][['Plant', 'Event', 'Count', 'Count_num']]
    df_inf = df_inf.rename(columns={'Count':'Influent_raw', 'Count_num':'Influent_used'})

    df_eff = get_effluent_df(microorganism)[['Plant', 'Event', 'Count', 'Count_num']]
    df_eff = df_eff.rename(columns={'Count':'Effluent_raw', 'Count_num':'Effluent_used'})

    df_paired = pd.merge(df_inf, df_eff, on=['Plant', 'Event'], how='inner')

    mdl = get_mdl(microorganism)
    df_paired = df_paired[df_paired['Influent_used'] > mdl]

    df_paired = df_paired.dropna(subset=['Influent_used', 'Effluent_used'])

    df_paired['LRV'] = np.log10(df_paired['Influent_used'] / df_paired['Effluent_used'])

    return df_paired

# Prepare results list
result_rows = []

for pathogen in pathogens:
    df_pathogen_pairs = get_lrv_pairs(pathogen)

    for indicator in indicators:
        df_indicator_pairs = get_lrv_pairs(indicator)

        # Merge pathogen and indicator LRVs on Plant + Event
        df_merged = pd.merge(
            df_pathogen_pairs[['Plant', 'Event', 'LRV']],
            df_indicator_pairs[['Plant', 'Event', 'LRV']],
            on=['Plant', 'Event'],
            suffixes=('_Pathogen', '_Indicator')
        )

        n_pairs = len(df_merged)
        if n_pairs < 3:
            spearman_rho = np.nan
            pearson_r = np.nan
        else:
            spearman_rho, _ = spearmanr(df_merged['LRV_Pathogen'], df_merged['LRV_Indicator'])
            pearson_r, _ = pearsonr(df_merged['LRV_Pathogen'], df_merged['LRV_Indicator'])

        result_rows.append({
            'Pathogen': pathogen,
            'Indicator': indicator,
            'n_pairs [-]': n_pairs,
            'Spearman_rho [-]': spearman_rho,
            'Pearson_r [-]': pearson_r
        })

# Create dataframe and save CSV
df_result = pd.DataFrame(result_rows)

output_filename = "correlation_indicator_pathogen_lrv.csv"
df_result.to_csv(output_filename, index=False)

print(output_filename)
print("correlation_table.py")