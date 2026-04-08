# filename: scatter_adenovirus_vs_indicators.py
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

# Organism lists for indicators and Lookup maps
indicators = [
    'Total coliform', 'Fecal coliform', 'E. coli',
    'Somatic coliphage', 'Male-specific coliphage', 'Aerobic endospores'
]

# Map IndicatorName to final Microorganism names (cover Adenovirus too)
indicator_map = {
    'totalcoliform': 'Total coliform',
    'fecalcoliform': 'Fecal coliform',
    'e.coli': 'E. coli',
    'somacoliphage': 'Somatic coliphage',
    'malespeccoliphage': 'Male-specific coliphage',
    'aerobicendospore': 'Aerobic endospores',
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

# MDL and Units dicts keyed by lowercase organism names
mdl_dict = {row['Organism'].lower(): row['MDL'] for _, row in df_mdl.iterrows()}

# Organism to MDL sheet Organism names map
microorganism_to_mdl_organism = {
    'Total coliform': 'Total coliform',
    'Fecal coliform': 'Fecal coliform',
    'E. coli': 'E. coli',
    'Somatic coliphage': 'Soma Coliphage',
    'Male-specific coliphage': 'Male Specific Coliphage',
    'Aerobic endospores': 'Aerobic endospores',
    'Adenovirus': 'Adenovirus',
}

# Map Microorganism to MDL and assign as column
def get_mdl(microorganism):
    mdl_org = microorganism_to_mdl_organism.get(microorganism)
    if mdl_org:
        return mdl_dict.get(mdl_org.lower(), np.nan)
    else:
        return np.nan

df_data['MDL'] = df_data['Microorganism'].apply(get_mdl)

# Replace BDL and 0 with MDL for calculations
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

# Sample type rules:
# Adenovirus effluent = Effluent grab
# Somatic and male-specific coliphage effluent = UF Effluent
# Others effluent = Effluent grab
uf_effluent_organisms = {'Somatic coliphage', 'Male-specific coliphage'}

# Helper to select effluent dataframe by microorganism
def get_effluent_df(microorganism):
    if microorganism == 'Adenovirus':
        return df_data[(df_data['SampleType'] == 'Effluent grab')]
    elif microorganism in uf_effluent_organisms:
        return df_data[(df_data['SampleType'] == 'UF Effluent')]
    else:
        return df_data[(df_data['SampleType'] == 'Effluent grab')]

# Influent dataframe is always Influent grab
df_influent = df_data[df_data['SampleType'] == 'Influent grab']

# Function to retrieve LRV pairs dataframe for given microorganism and plant
def get_LRV_pairs(microorganism, plant=None):
    # Influent filtered
    df_inf = df_influent[df_influent['Microorganism'] == microorganism]
    if plant:
        df_inf = df_inf[df_inf['Plant'] == plant]
    # Effluent filtered
    df_eff = get_effluent_df(microorganism)
    df_eff = df_eff[df_eff['Microorganism'] == microorganism]
    if plant:
        df_eff = df_eff[df_eff['Plant'] == plant]

    # Select relevant columns and rename counts
    df_inf_sel = df_inf[['Plant', 'Event', 'Count', 'Count_num']].rename(columns={'Count': 'Influent_raw', 'Count_num': 'Influent_used'})
    df_eff_sel = df_eff[['Plant', 'Event', 'Count', 'Count_num']].rename(columns={'Count': 'Effluent_raw', 'Count_num': 'Effluent_used'})

    # Merge influent and effluent by Plant and Event
    df_paired = pd.merge(df_inf_sel, df_eff_sel, on=['Plant', 'Event'], how='inner')

    # Filter pairs where influent_used > MDL to exclude low influent values
    mdl = get_mdl(microorganism)
    df_paired = df_paired[df_paired['Influent_used'] > mdl]

    # Drop NaNs in used counts
    df_paired = df_paired.dropna(subset=['Influent_used', 'Effluent_used'])

    # Calculate LRV
    df_paired['LRV'] = np.log10(df_paired['Influent_used'] / df_paired['Effluent_used'])

    return df_paired

# Get Adenovirus pairs (all plants pooled)
df_adenovirus_pairs = get_LRV_pairs('Adenovirus')

# We'll gather dataframes for each indicator organism with paired Adenovirus
list_scatter_data = []

for indicator in indicators:
    df_ind_pairs = get_LRV_pairs(indicator)

    # Merge Adenovirus with indicator pairs on Plant + Event
    # Keep only pairs where both LRVs exist
    df_merge = pd.merge(
        df_ind_pairs[['Plant','Event','LRV']],
        df_adenovirus_pairs[['Plant','Event','LRV']],
        on=['Plant','Event'],
        suffixes=('_Indicator', '_Adenovirus')
    )

    df_merge['Indicator'] = indicator

    list_scatter_data.append(df_merge)

# Concatenate data for all indicator panels
df_all = pd.concat(list_scatter_data, ignore_index=True)

# Rename columns for output clarity
df_all = df_all.rename(columns={
    'LRV_Indicator': 'Indicator_LRV [log10]',
    'LRV_Adenovirus': 'Adenovirus_LRV [log10]'
})

# Plotly multi-panel scatter plot
plants_order = sorted(df_data['Plant'].dropna().unique())
indicators_ordered = indicators

panel_rows = 2
panel_cols = 3

# Remove global y_title and x_title from make_subplots
fig = make_subplots(
    rows=panel_rows, cols=panel_cols,
    subplot_titles=indicators_ordered,
    shared_xaxes=False,
    shared_yaxes=True,
    vertical_spacing=0.15,
    horizontal_spacing=0.12,
)

# Colors and symbols for detection based on Adenovirus effluent raw detection (we do detection plot only on effluent for Adenovirus)
# Effluent detection flag for Adenovirus calculated from df_adenovirus_pairs original raw counts
# We'll create dict keyed by Plant+Event tuple for Adenovirus detection flags
def get_effluent_detect_flag_from_raw_count(val):
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

# Get Adenovirus raw effluent data for plant+event detection lookup
df_adenovirus_eff_raw = df_data[(df_data['Microorganism']=='Adenovirus')&(df_data['SampleType']=='Effluent grab')]
df_adenovirus_eff_raw = df_adenovirus_eff_raw.set_index(['Plant','Event'])
adenovirus_detect_map = df_adenovirus_eff_raw['Count'].apply(get_effluent_detect_flag_from_raw_count).to_dict()

# Marker config
color_map = {'Non-detect': 'blue', 'Detected': 'red'}
symbol_map = {'Non-detect': 'circle-open', 'Detected': 'x'}

for i, indicator in enumerate(indicators_ordered):
    row = (i // panel_cols) + 1
    col = (i % panel_cols) + 1

    df_panel = df_all[df_all['Indicator'] == indicator]

    # Add scatter points for each pair colored by Adenovirus effluent detection
    for detect_flag in ['Non-detect', 'Detected']:
        df_sub = df_panel.copy()
        # Map plant,event to detection flag; only keep matching flags for plotting
        df_sub['Effluent_detected_flag'] = df_sub.apply(
            lambda r: adenovirus_detect_map.get((r['Plant'], r['Event']), 'Unknown'), axis=1)
        df_sub = df_sub[df_sub['Effluent_detected_flag'] == detect_flag]

        fig.add_trace(go.Scatter(
            x=df_sub['Indicator_LRV [log10]'],
            y=df_sub['Adenovirus_LRV [log10]'],
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

    # Configure axes
    # Dynamic range based on data range for plot
    x_vals = df_panel['Indicator_LRV [log10]']
    y_vals = df_panel['Adenovirus_LRV [log10]']
    if not x_vals.empty and not y_vals.empty:
        x_min, x_max = x_vals.quantile(0.02), x_vals.quantile(0.98)
        y_min, y_max = y_vals.quantile(0.02), y_vals.quantile(0.98)
        extra = 0.5
        fig.update_xaxes(range=[x_min - extra, x_max + extra], row=row, col=col, zeroline=True, zerolinewidth=1, zerolinecolor='lightgray', automargin=True)
        fig.update_yaxes(range=[y_min - extra, y_max + extra], row=row, col=col, zeroline=True, zerolinewidth=1, zerolinecolor='lightgray', automargin=True)

    # Set y-axis title only on first column of panels (left side), with automargin to prevent label clipping
    if col == 1:
        fig.update_yaxes(title_text='Adenovirus LRV [log10]', row=row, col=col, automargin=True)
    else:
        # Remove y-axis title for other columns
        fig.update_yaxes(title_text=None, row=row, col=col)

    # Set x-axis title only on bottom row panels (row=2), with automargin and standoff for spacing
    if row == panel_rows:
        fig.update_xaxes(title_text='Indicator LRV [log10]', row=row, col=col, automargin=True, title_standoff=15)
    else:
        # Remove x-axis titles for top row panels
        fig.update_xaxes(title_text=None, row=row, col=col)

fig.update_layout(
    font=dict(size=16),
    height=700,
    width=1200,
    margin=dict(t=50, b=80, l=80, r=40),  # increased bottom margin for x-axis titles clearance
    template='plotly_white',
    legend=dict(
        title='Adenovirus Effluent Detection'
    )
)

# Save outputs to same filenames as original
base_filename = "scatter_adenovirus_vs_indicators"
html_file = base_filename + ".html"
png_file = base_filename + ".png"
csv_file = base_filename + ".csv"

fig.write_html(html_file, include_plotlyjs='cdn')
fig.write_image(png_file, width=1200, height=700)
df_all.to_csv(csv_file, index=False)

print(csv_file)
print(html_file)
print(png_file)
print("scatter_adenovirus_vs_indicators.py")