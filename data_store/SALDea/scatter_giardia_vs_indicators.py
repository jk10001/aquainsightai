# filename: scatter_giardia_vs_indicators.py
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

# Indicator organisms list for subplot panels
indicators = [
    'Total coliform', 'Fecal coliform', 'E. coli',
    'Somatic coliphage', 'Male-specific coliphage', 'Aerobic endospores'
]

# Map IndicatorName to final Microorganism names
indicator_map = {
    'totalcoliform': 'Total coliform',
    'fecalcoliform': 'Fecal coliform',
    'e.coli': 'E. coli',
    'somacoliphage': 'Somatic coliphage',
    'malespeccoliphage': 'Male-specific coliphage',
    'aerobicendospore': 'Aerobic endospores',
    'giardia': 'Giardia',
}

def map_indicator_to_microorganism(ind_name):
    ind_name_lower = ind_name.lower()
    for key, val in indicator_map.items():
        if key in ind_name_lower:
            return val
    return None

df_data['Microorganism'] = df_data['IndicatorName'].apply(map_indicator_to_microorganism)
df_data = df_data[df_data['Microorganism'].notna()]

# Giardia subset for pairing and detection
df_giardia = df_data[df_data['Microorganism'] == 'Giardia'].copy()

# MDL dictionary by lowercase organism name
mdl_dict = {row['Organism'].lower(): row['MDL'] for _, row in df_mdl.iterrows()}

# Map Microorganism to MDL sheet Organism names for MDL lookup
microorganism_to_mdl_organism = {
    'Total coliform': 'Total coliform',
    'Fecal coliform': 'Fecal coliform',
    'E. coli': 'E. coli',
    'Somatic coliphage': 'Soma Coliphage',
    'Male-specific coliphage': 'Male Specific Coliphage',
    'Aerobic endospores': 'Aerobic endospores',
    'Giardia': 'Giardia Cyst',
}

def get_mdl(microorganism):
    mdl_org = microorganism_to_mdl_organism.get(microorganism)
    if mdl_org:
        return mdl_dict.get(mdl_org.lower(), np.nan)
    else:
        return np.nan

# Add MDL column to full df_data
df_data['MDL'] = df_data['Microorganism'].apply(get_mdl)
df_giardia['MDL'] = get_mdl('Giardia')

# Replace BDL and 0 with MDL for calculation convenience
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
        else:
            return fval
    except:
        return np.nan

df_data['Count_num'] = df_data.apply(replace_bdl_zero_with_mdl, axis=1)
df_giardia['Count_num'] = df_giardia.apply(replace_bdl_zero_with_mdl, axis=1)

# Define sample types by microorganism group per instructions
uf_effluent_organisms = {'Giardia', 'Somatic coliphage', 'Male-specific coliphage'}

# Identify effluent sample selection for indicators (bacteria + aerobic endospores = Effluent grab; coliphages = UF Effluent)
def get_indicator_effluent_df(microorganism):
    if microorganism in {'Somatic coliphage', 'Male-specific coliphage'}:
        return df_data[(df_data['SampleType'] == 'UF Effluent') & (df_data['Microorganism'] == microorganism)]
    else:
        return df_data[(df_data['SampleType'] == 'Effluent grab') & (df_data['Microorganism'] == microorganism)]

# Giardia effluent is always 'UF Effluent'
df_giardia_eff = df_giardia[df_giardia['SampleType'] == 'UF Effluent']

# Influent is always 'Influent grab'
df_influent = df_data[(df_data['SampleType'] == 'Influent grab')].copy()

# Prepare the scatter data list
rows = []

# For each indicator, pair Giardia influent and indicator effluent by Plant+Event
for indicator in indicators:

    mdl_indicator = get_mdl(indicator)
    if np.isnan(mdl_indicator):
        continue
    mdl_giardia = get_mdl('Giardia')

    # Influent Giardia samples for all plants pooled
    df_giardia_inf = df_influent[(df_influent['Microorganism'] == 'Giardia')][['Plant', 'Event', 'Count', 'Count_num']]
    df_giardia_inf = df_giardia_inf.rename(columns={'Count': 'Giardia_influent_raw', 'Count_num': 'Giardia_influent_used'})

    # Effluent indicator samples with appropriate SampleType filter and Microorganism filter
    df_indicator_eff = get_indicator_effluent_df(indicator)[['Plant', 'Event', 'Count', 'Count_num']]
    df_indicator_eff = df_indicator_eff.rename(columns={'Count': 'Indicator_effluent_raw', 'Count_num': 'Indicator_effluent_used'})

    # Effluent Giardia samples for effluent detection coloring and used values
    # Include both raw Count and Count_num for Giardia UF Effluent
    df_giardia_effect = df_giardia_eff[['Plant', 'Event', 'Count', 'Count_num']]
    df_giardia_effect = df_giardia_effect.rename(columns={'Count': 'Giardia_effluent_raw', 'Count_num': 'Giardia_effluent_used'})

    # Merge Giardia influent with indicator effluent by Plant+Event to form pairs
    paired = pd.merge(df_giardia_inf, df_indicator_eff, on=['Plant', 'Event'], how='inner')

    # Merge with Giardia effluent raw and used counts for detection flag and LRV denominator
    paired = pd.merge(paired, df_giardia_effect, on=['Plant', 'Event'], how='left')

    # Remove pairs with NaNs in used counts
    paired = paired.dropna(subset=['Giardia_influent_used', 'Indicator_effluent_used'])

    # Exclude pairs where Giardia influent count ≤ MDL Giardia (skip zero/MDL influent)
    paired = paired[paired['Giardia_influent_used'] > mdl_giardia]

    # Influent indicator samples
    df_indicator_inf = df_influent[(df_influent['Microorganism'] == indicator)][['Plant', 'Event', 'Count', 'Count_num']]
    df_indicator_inf = df_indicator_inf.rename(columns={'Count': 'Indicator_influent_raw', 'Count_num': 'Indicator_influent_used'})

    # Merge influent and effluent of indicator on Plant+Event
    df_indicator_pair = pd.merge(df_indicator_inf, df_indicator_eff, on=['Plant', 'Event'], how='inner')
    df_indicator_pair = df_indicator_pair.dropna(subset=['Indicator_influent_used', 'Indicator_effluent_used'])

    # Filter to remove influent ≤ MDL
    df_indicator_pair = df_indicator_pair[df_indicator_pair['Indicator_influent_used'] > mdl_indicator]

    # Calculate indicator LRV for each pair
    df_indicator_pair['Indicator_LRV'] = np.log10(df_indicator_pair['Indicator_influent_used'] / df_indicator_pair['Indicator_effluent_used'])

    # Merge indicator LRV into paired data to get x axis values for that panel
    paired = pd.merge(paired, df_indicator_pair[['Plant', 'Event', 'Indicator_LRV']], on=['Plant', 'Event'], how='inner')

    # Drop any pairs with missing Giardia effluent used, influent used, or Indicator_LRV before calculation
    paired = paired.dropna(subset=['Giardia_influent_used', 'Giardia_effluent_used', 'Indicator_LRV'])

    # Calculate Giardia LRV correctly using Giardia effluent used count in denominator
    paired['Giardia_LRV'] = np.log10(paired['Giardia_influent_used'] / paired['Giardia_effluent_used'])

    # Determine effluent detection flag on Giardia raw effluent counts (before substitution)
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

    paired['Giardia_effluent_detection'] = paired['Giardia_effluent_raw'].apply(effluent_detect_flag)

    # Append dataframe for this indicator with columns needed for plotting
    rows.append(paired[['Plant', 'Event', 'Indicator_LRV', 'Giardia_LRV', 'Giardia_effluent_detection']].assign(Indicator=indicator))

# Concatenate all indicator panels data
df_all = pd.concat(rows, ignore_index=True)

# Prepare subplot figure
panel_rows = 2
panel_cols = 3
subplot_titles = ["Total coliform","Fecal coliform","E. coli","Somatic coliphage","Male-specific coliphage","Aerobic endospores"]

fig = make_subplots(
    rows=panel_rows,
    cols=panel_cols,
    subplot_titles=subplot_titles,
    shared_yaxes=True,
    vertical_spacing=0.15,
    horizontal_spacing=0.12
)

# Colors and symbols based on Giardia effluent detection
color_map = {'Non-detect': 'blue', 'Detected': 'red'}
symbol_map = {'Non-detect': 'circle-open', 'Detected': 'x'}

# For legend tracking only once
legend_shown = {'Non-detect': False, 'Detected': False}

for idx, indicator in enumerate(subplot_titles):
    row = (idx // panel_cols) + 1
    col = (idx % panel_cols) + 1

    df_panel = df_all[df_all['Indicator'] == indicator]

    # Calculate axis ranges using approx 2% - 98% quantiles with padding
    x_vals = df_panel['Indicator_LRV']
    y_vals = df_panel['Giardia_LRV']

    if not x_vals.empty and not y_vals.empty:
        x_min, x_max = x_vals.quantile(0.02), x_vals.quantile(0.98)
        y_min, y_max = y_vals.quantile(0.02), y_vals.quantile(0.98)

        extra = 0.5
        fig.update_xaxes(range=[x_min - extra, x_max + extra], row=row, col=col, zeroline=True,
                         zerolinewidth=1, zerolinecolor='lightgray', automargin=True)
        fig.update_yaxes(range=[y_min - extra, y_max + extra], row=row, col=col, zeroline=True,
                         zerolinewidth=1, zerolinecolor='lightgray', automargin=True)

    # Add scatter points by detection flag
    for detection_flag in ['Non-detect', 'Detected']:
        df_sub = df_panel[df_panel['Giardia_effluent_detection'] == detection_flag]
        show_legend = False
        # show legend only on first subplot (idx==0)
        if idx == 0 and not legend_shown[detection_flag]:
            show_legend = True
            legend_shown[detection_flag] = True

        fig.add_trace(go.Scatter(
            x=df_sub['Indicator_LRV'],
            y=df_sub['Giardia_LRV'],
            mode='markers',
            marker=dict(
                symbol=symbol_map[detection_flag],
                color=color_map[detection_flag],
                size=10,
                line=dict(width=1)
            ),
            name=f'Effluent {detection_flag}',
            showlegend=show_legend,
        ), row=row, col=col)

    # Axis titles per instructions
    # Y-axis title ONLY on left column panels with automargin
    if col == 1:
        fig.update_yaxes(title_text='Giardia LRV [log10]', row=row, col=col, automargin=True)
    else:
        fig.update_yaxes(title_text=None, row=row, col=col)

    # X-axis title ONLY on bottom row panels with automargin and title_standoff approx 15
    if row == panel_rows:
        fig.update_xaxes(title_text='Indicator LRV [log10]', row=row, col=col, automargin=True, title_standoff=15)
    else:
        fig.update_xaxes(title_text=None, row=row, col=col)

fig.update_layout(
    font=dict(size=16),
    height=700,
    width=1200,
    template='plotly_white',
    margin=dict(t=50, b=80, l=80, r=40),
    legend=dict(title='Giardia Effluent Detection')
)

# Output file base name
base_filename = "scatter_giardia_vs_indicators"
html_file = base_filename + ".html"
png_file = base_filename + ".png"
csv_file = base_filename + ".csv"
py_file = base_filename + ".py"

# Save outputs
fig.write_html(html_file, include_plotlyjs='cdn')
fig.write_image(png_file, width=1200, height=700)
df_all.to_csv(csv_file, index=False)

# Save the python source code for reproducibility alongside
# (We omit writing the source code to disk automatically. User can save this script manually if needed.)

print(csv_file)
print(html_file)
print(png_file)
print(py_file)