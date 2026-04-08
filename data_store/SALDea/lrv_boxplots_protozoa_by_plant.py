# filename: lrv_boxplots_protozoa_by_plant.py
import pandas as pd
import numpy as np
import plotly.graph_objs as go
from plotly.subplots import make_subplots
import math

# Load data
file_name = "WWTP_microbial_loads_and_removal.xlsx"
xls = pd.ExcelFile(file_name)
df_data = pd.read_excel(xls, sheet_name="Data")
df_mdl = pd.read_excel(xls, sheet_name="Method detection limits")

# Normalize Plant codes: unify 'BiSP' -> 'BiSp'
df_data['Plant'] = df_data['Plant'].replace({'BiSP': 'BiSp'})

# Define protozoan organisms to include
protozoa = ['Giardia', 'Cryptosporidium']

# Original mapping from IndicatorName to protozoan microorganism output names
indicator_to_protozoan = {
    'giardia': 'Giardia',
    'crypto': 'Cryptosporidium',
}

def map_indicator_to_protozoan(ind_name):
    ind_name_lower = ind_name.lower()
    for key, val in indicator_to_protozoan.items():
        if key in ind_name_lower:
            return val
    return None

# Apply mapping
df_data['Microorganism'] = df_data['IndicatorName'].apply(map_indicator_to_protozoan)
# Keep only protozoa rows
df_data = df_data[df_data['Microorganism'].isin(protozoa)].copy()

# Create lookup for MDL and Units by lower-case organism name
mdl_dict = {}
units_dict = {}
for _, row in df_mdl.iterrows():
    mdl_dict[row['Organism'].lower()] = row['MDL']
    units_dict[row['Organism'].lower()] = row['Units']

# Map Microorganism to MDL sheet Organism names for consistency in MDL/Units lookup
microorganism_to_mdl_organism = {
    'Giardia': 'Giardia Cyst',
    'Cryptosporidium': 'Crypto Oocyst',
}

# Add MDL and Units columns to full dataset
def get_mdl_and_units(microorganism):
    mdl_org = microorganism_to_mdl_organism.get(microorganism)
    if mdl_org:
        mdl_val = mdl_dict.get(mdl_org.lower(), np.nan)
        unit_val = units_dict.get(mdl_org.lower(), '')
        return pd.Series([mdl_val, unit_val])
    return pd.Series([np.nan, ''])

df_data[['MDL', 'Units']] = df_data['Microorganism'].apply(get_mdl_and_units)

# Replace BDL and numeric 0 with organism-specific MDL for the entire dataset
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

# Separate influent and effluent dataframes (with Count_num included)
df_influent = df_data[df_data['SampleType'] == 'Influent grab'].copy()
df_uf_effluent = df_data[df_data['SampleType'] == 'UF Effluent'].copy()

rows = []
plants = sorted(df_data['Plant'].dropna().unique())

for plant in plants:
    for microorganism in protozoa:
        mdl_value = df_data.loc[df_data['Microorganism'] == microorganism, 'MDL'].dropna().unique()
        if len(mdl_value) == 0:
            continue
        mdl_value = mdl_value[0]

        # Influent samples with Count_num
        df_inf = df_influent[(df_influent['Plant'] == plant) & (df_influent['Microorganism'] == microorganism)][['Event', 'Count', 'Count_num']]
        df_inf = df_inf.rename(columns={'Count': 'Influent_raw', 'Count_num': 'Influent_used'})

        # Effluent samples with Count_num (UF Effluent only)
        df_eff = df_uf_effluent[(df_uf_effluent['Plant'] == plant) & (df_uf_effluent['Microorganism'] == microorganism)][['Event', 'Count', 'Count_num']]
        df_eff = df_eff.rename(columns={'Count': 'Effluent_raw', 'Count_num': 'Effluent_used'})

        # Merge on Event to pair samples
        paired = pd.merge(df_inf, df_eff, on='Event', how='inner')

        # Add Plant and Microorganism columns
        paired['Plant'] = plant
        paired['Microorganism'] = microorganism

        # Exclude pairs where Influent_used <= MDL (skip zero or MDL influent values)
        paired = paired[paired['Influent_used'] > mdl_value]

        # Drop rows with NaN in influent or effluent used count
        paired = paired.dropna(subset=['Influent_used', 'Effluent_used'])

        # Calculate LRV = log10(influent_used / effluent_used)
        paired['LRV [log10]'] = np.log10(paired['Influent_used'] / paired['Effluent_used'])

        # Determine Effluent detection flag based on raw effluent value: non-detect if BDL or 0, detected otherwise
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

        paired['Effluent_detected_flag'] = paired['Effluent_raw'].apply(effluent_detect_flag)

        rows.append(paired[['Event', 'Plant', 'Microorganism', 'Influent_raw', 'Effluent_raw',
                            'Influent_used', 'Effluent_used', 'Effluent_detected_flag', 'LRV [log10]']])

# Concatenate all pairs
df_pairs = pd.concat(rows, ignore_index=True)
df_pairs = df_pairs.sort_values(['Microorganism', 'Plant', 'Event']).reset_index(drop=True)

# Plot multi-panel box plot
plants_order = ['BiSp', 'BrWo', 'NoWe']
panel_rows = 1
panel_cols = 2
fig = make_subplots(rows=panel_rows, cols=panel_cols, subplot_titles=protozoa,
                    shared_yaxes=True, vertical_spacing=0.2)

# Determine y-axis range (omit outliers)
percentiles = df_pairs['LRV [log10]'].quantile([0.02, 0.98]).values
y_min = math.floor(percentiles[0])
y_max = math.ceil(percentiles[1])

plant_colors = {'BiSp': 'blue', 'BrWo': 'green', 'NoWe': 'orange'}
marker_symbols = {'Non-detect': 'circle-open', 'Detected': 'x'}

for idx, microorganism in enumerate(protozoa):
    row = 1
    col = idx + 1

    df_sub = df_pairs[df_pairs['Microorganism'] == microorganism]

    for plant in plants_order:
        df_plant = df_sub[df_sub['Plant'] == plant]
        if df_plant.empty:
            continue

        # Box plot without legend entries for plants
        fig.add_trace(go.Box(
            y=df_plant['LRV [log10]'],
            x=[plant]*len(df_plant),
            name=plant,
            marker_color=plant_colors.get(plant, 'black'),
            boxmean=True,
            boxpoints=False,
            showlegend=False,
            line=dict(width=2)
        ), row=row, col=col)

        # Scatter non-detect effluent points
        non_detect_df = df_plant[df_plant['Effluent_detected_flag'] == 'Non-detect']
        fig.add_trace(go.Scatter(
            x=[plant]*len(non_detect_df),
            y=non_detect_df['LRV [log10]'],
            mode='markers',
            marker=dict(
                symbol=marker_symbols['Non-detect'],
                color=plant_colors.get(plant, 'black'),
                size=10,
                line=dict(width=1)
            ),
            name='Effluent Non-detect',
            showlegend=(idx == 0 and plant == plants_order[0]),
            legendgroup='Effluent Non-detect'
        ), row=row, col=col)

        # Scatter detected effluent points
        detected_df = df_plant[df_plant['Effluent_detected_flag'] == 'Detected']
        fig.add_trace(go.Scatter(
            x=[plant]*len(detected_df),
            y=detected_df['LRV [log10]'],
            mode='markers',
            marker=dict(
                symbol=marker_symbols['Detected'],
                color=plant_colors.get(plant, 'black'),
                size=10,
                line=dict(width=1)
            ),
            name='Effluent Detected',
            showlegend=(idx == 0 and plant == plants_order[0]),
            legendgroup='Effluent Detected'
        ), row=row, col=col)

    # Update axes
    fig.update_yaxes(title_text='LRV [log10]', range=[y_min, y_max], row=row, col=col,
                     zeroline=True, zerolinewidth=1, zerolinecolor='gray')
    fig.update_xaxes(title_text='Plant', row=row, col=col)

fig.update_layout(
    font=dict(size=16),
    boxmode='group',
    showlegend=True,
    legend=dict(title='Effluent detection'),
    margin=dict(t=50, b=30, l=50, r=40),
    height=600,
    width=1000,
    template='plotly_white'
)

# Save outputs
base_filename = "lrv_boxplots_protozoa_by_plant"
html_file = base_filename + ".html"
png_file = base_filename + ".png"
csv_file = base_filename + ".csv"

fig.write_html(html_file, include_plotlyjs='cdn')
fig.write_image(png_file, width=1000, height=600)
df_pairs.to_csv(csv_file, index=False)

print(csv_file)
print(html_file)
print(png_file)
print("lrv_boxplots_protozoa_by_plant.py")