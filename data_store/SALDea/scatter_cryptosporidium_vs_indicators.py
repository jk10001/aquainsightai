# filename: scatter_cryptosporidium_vs_indicators.py
import pandas as pd
import numpy as np
import plotly.graph_objs as go
from plotly.subplots import make_subplots
import math

file_name = "WWTP_microbial_loads_and_removal.xlsx"
xls = pd.ExcelFile(file_name)
df_data = pd.read_excel(xls, sheet_name="Data")
df_mdl = pd.read_excel(xls, sheet_name="Method detection limits")

# Normalize Plant codes
df_data['Plant'] = df_data['Plant'].replace({'BiSP': 'BiSp'})

# Indicator organisms list (panels)
indicators = [
    'Total coliform', 'Fecal coliform', 'E. coli',
    'Somatic coliphage', 'Male-specific coliphage', 'Aerobic endospores'
]

# Map IndicatorName to final Microorganism names (including Cryptosporidium)
indicator_map = {
    'totalcoliform': 'Total coliform',
    'fecalcoliform': 'Fecal coliform',
    'e.coli': 'E. coli',
    'somacoliphage': 'Somatic coliphage',
    'malespeccoliphage': 'Male-specific coliphage',
    'aerobicendospore': 'Aerobic endospores',
    'crypto': 'Cryptosporidium',
}

def map_indicator_to_microorganism(ind_name):
    ind_name_lower = ind_name.lower()
    for key, val in indicator_map.items():
        if key in ind_name_lower:
            return val
    return None

df_data['Microorganism'] = df_data['IndicatorName'].apply(map_indicator_to_microorganism)
df_data = df_data[df_data['Microorganism'].notna()]

# MDL dictionary by lowercase organism name
mdl_dict = {row['Organism'].lower(): row['MDL'] for _, row in df_mdl.iterrows()}

# Microorganism to MDL organism name mapping
microorganism_to_mdl_organism = {
    'Total coliform': 'Total coliform',
    'Fecal coliform': 'Fecal coliform',
    'E. coli': 'E. coli',
    'Somatic coliphage': 'Soma Coliphage',
    'Male-specific coliphage': 'Male Specific Coliphage',
    'Aerobic endospores': 'Aerobic endospores',
    'Cryptosporidium': 'Crypto Oocyst',
}

# Function to get MDL for microorganism
def get_mdl(microorganism):
    mdl_org = microorganism_to_mdl_organism.get(microorganism)
    if mdl_org:
        return mdl_dict.get(mdl_org.lower(), np.nan)
    else:
        return np.nan

df_data['MDL'] = df_data['Microorganism'].apply(get_mdl)

# Replace BDL and numeric 0 with MDL for calculations
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

# Influent is always Influent grab
df_influent = df_data[df_data['SampleType'] == 'Influent grab']

# Effluent selection rules:
# Cryptosporidium and coliphages effluent = UF Effluent
# Others = Effluent grab
uf_effluent_organisms = {'Cryptosporidium', 'Somatic coliphage', 'Male-specific coliphage'}

def get_effluent_df(microorganism):
    if microorganism in uf_effluent_organisms:
        return df_data[df_data['SampleType'] == 'UF Effluent']
    else:
        return df_data[df_data['SampleType'] == 'Effluent grab']

# Function to get LRV pairs for microorganism and plant (or all plants)
def get_LRV_pairs(microorganism, plant=None):
    df_inf = df_influent[df_influent['Microorganism'] == microorganism]
    if plant:
        df_inf = df_inf[df_inf['Plant'] == plant]
    df_eff = get_effluent_df(microorganism)
    df_eff = df_eff[df_eff['Microorganism'] == microorganism]
    if plant:
        df_eff = df_eff[df_eff['Plant'] == plant]

    df_inf_sel = df_inf[['Plant','Event','Count','Count_num']].rename(columns={'Count':'Influent_raw','Count_num':'Influent_used'})
    df_eff_sel = df_eff[['Plant','Event','Count','Count_num']].rename(columns={'Count':'Effluent_raw','Count_num':'Effluent_used'})

    df_paired = pd.merge(df_inf_sel, df_eff_sel, on=['Plant','Event'], how='inner')

    mdl = get_mdl(microorganism)
    df_paired = df_paired[df_paired['Influent_used'] > mdl]

    df_paired = df_paired.dropna(subset=['Influent_used', 'Effluent_used'])
    df_paired['LRV'] = np.log10(df_paired['Influent_used'] / df_paired['Effluent_used'])

    return df_paired

# Get Cryptosporidium pairs all plants pooled
df_crypto_pairs = get_LRV_pairs('Cryptosporidium')

list_scatter_data = []

for indicator in indicators:
    df_ind_pairs = get_LRV_pairs(indicator)

    df_merge = pd.merge(
        df_ind_pairs[['Plant','Event','LRV']],
        df_crypto_pairs[['Plant','Event','LRV']],
        on=['Plant','Event'],
        suffixes=('_Indicator','_Cryptosporidium')
    )

    df_merge['Indicator'] = indicator

    list_scatter_data.append(df_merge)

df_all = pd.concat(list_scatter_data, ignore_index=True)

df_all = df_all.rename(columns={
    'LRV_Indicator': 'Indicator_LRV [log10]',
    'LRV_Cryptosporidium': 'Cryptosporidium_LRV [log10]'
})

# Determine Cryptosporidium effluent detection flag for coloring
def effluent_detect_flag(val):
    if pd.isna(val):
        return 'Unknown'
    val_str = str(val).strip().upper()
    if val_str == 'BDL':
        return 'Non-detect'
    try:
        fval = float(val)
        if fval == 0:
            return 'Non-detect'
        else:
            return 'Detected'
    except:
        return 'Unknown'

df_crypto_eff_raw = df_data[(df_data['Microorganism']=='Cryptosporidium') & (df_data['SampleType']=='UF Effluent')]
df_crypto_eff_raw = df_crypto_eff_raw.set_index(['Plant','Event'])
crypto_detect_map = df_crypto_eff_raw['Count'].apply(effluent_detect_flag).to_dict()

import plotly.graph_objs as go
from plotly.subplots import make_subplots

plants_order = sorted(df_data['Plant'].dropna().unique())
indicators_ordered = indicators

panel_rows = 2
panel_cols = 3

# Remove global x_title and y_title here as per instructions
fig = make_subplots(
    rows=panel_rows, cols=panel_cols,
    subplot_titles=indicators_ordered,
    shared_xaxes=False,
    shared_yaxes=True,
    vertical_spacing=0.15,
    horizontal_spacing=0.12,
    # Removed x_title and y_title to fix overlapping labels as instructed
)

# Marker config and colors based on Cryptosporidium effluent detection
color_map = {'Non-detect': 'blue', 'Detected': 'red'}
symbol_map = {'Non-detect': 'circle-open', 'Detected': 'x'}

for i, indicator in enumerate(indicators_ordered):
    row = (i // panel_cols) + 1
    col = (i % panel_cols) + 1

    df_panel = df_all[df_all['Indicator'] == indicator]

    for detect_flag in ['Non-detect', 'Detected']:
        df_sub = df_panel.copy()
        df_sub['Effluent_detected_flag'] = df_sub.apply(
            lambda r: crypto_detect_map.get((r['Plant'], r['Event']), 'Unknown'), axis=1)
        df_sub = df_sub[df_sub['Effluent_detected_flag'] == detect_flag]

        fig.add_trace(go.Scatter(
            x=df_sub['Indicator_LRV [log10]'],
            y=df_sub['Cryptosporidium_LRV [log10]'],
            mode='markers',
            marker=dict(
                symbol=symbol_map[detect_flag],
                color=color_map[detect_flag],
                size=10,
                line=dict(width=1)
            ),
            name=f'Effluent {detect_flag}',
            showlegend=True if i==0 else False
        ), row=row, col=col)

    # Configure axes ranges based on quantiles for each panel
    x_vals = df_panel['Indicator_LRV [log10]']
    y_vals = df_panel['Cryptosporidium_LRV [log10]']
    if not x_vals.empty and not y_vals.empty:
        x_min, x_max = x_vals.quantile(0.02), x_vals.quantile(0.98)
        y_min, y_max = y_vals.quantile(0.02), y_vals.quantile(0.98)
        extra = 0.5
        fig.update_xaxes(
            range=[x_min - extra, x_max + extra],
            row=row, col=col,
            zeroline=True, zerolinewidth=1, zerolinecolor='lightgray',
            automargin=True
        )
        fig.update_yaxes(
            range=[y_min - extra, y_max + extra],
            row=row, col=col,
            zeroline=True, zerolinewidth=1, zerolinecolor='lightgray',
            automargin=True
        )

    # Set y-axis title only for first column panels
    if col == 1:
        fig.update_yaxes(title_text='Cryptosporidium LRV [log10]', row=row, col=col, automargin=True)
    else:
        fig.update_yaxes(title_text=None, row=row, col=col)

    # Set x-axis title only for bottom row panels
    if row == panel_rows:
        fig.update_xaxes(
            title_text='Indicator LRV [log10]', row=row, col=col,
            automargin=True,
            title_standoff=15
        )
    else:
        fig.update_xaxes(title_text=None, row=row, col=col)

fig.update_layout(
    font=dict(size=16),
    height=700,
    width=1200,
    margin=dict(t=50, b=80, l=80, r=40),  # increased bottom margin b=80 for axis titles
    template='plotly_white',
    legend=dict(title='Cryptosporidium Effluent Detection')
)

# Save outputs
base_filename = "scatter_cryptosporidium_vs_indicators"
html_file = base_filename + ".html"
png_file = base_filename + ".png"
csv_file = base_filename + ".csv"

fig.write_html(html_file, include_plotlyjs='cdn')
fig.write_image(png_file, width=1200, height=700)
df_all.to_csv(csv_file, index=False)

print(csv_file)
print(html_file)
print(png_file)
print("scatter_cryptosporidium_vs_indicators.py")