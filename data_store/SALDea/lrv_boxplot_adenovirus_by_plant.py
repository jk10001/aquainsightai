# filename: lrv_boxplot_adenovirus_by_plant.py
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

# Filter only Adenovirus rows based on IndicatorName
def is_adenovirus(ind_name):
    return 'adenovirus' in ind_name.lower()

df_data = df_data[df_data['IndicatorName'].apply(is_adenovirus)].copy()

# Map Microorganism for consistency
df_data['Microorganism'] = 'Adenovirus'

# Create lookup for MDL and Units by Organism name (lowercase keys)
mdl_dict = {row['Organism'].lower(): row['MDL'] for _, row in df_mdl.iterrows()}
units_dict = {row['Organism'].lower(): row['Units'] for _, row in df_mdl.iterrows()}

# Get MDL and Units for Adenovirus
mdlv = mdl_dict.get('adenovirus', np.nan)
unitsv = units_dict.get('adenovirus', '')

# Assign MDL and Units columns
df_data['MDL'] = mdlv
df_data['Units'] = unitsv

# Define sample types for influent and effluent for Adenovirus
df_influent = df_data[df_data['SampleType'] == 'Influent grab'].copy()
df_effluent = df_data[df_data['SampleType'] == 'Effluent grab'].copy()

# Replace BDL and zero with MDL for count_num used in calculations
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
# Update influent and effluent data to include this column
df_influent = df_influent.copy()
df_influent['Count_num'] = df_influent.apply(replace_bdl_zero_with_mdl, axis=1)
df_effluent = df_effluent.copy()
df_effluent['Count_num'] = df_effluent.apply(replace_bdl_zero_with_mdl, axis=1)

# Prepare dataframe to store paired data for plotting
rows = []

plants = sorted(df_data['Plant'].dropna().unique())

for plant in plants:
    # Influent samples for plant
    df_inf = df_influent[df_influent['Plant'] == plant][['Event', 'Count', 'Count_num']]
    df_inf = df_inf.rename(columns={'Count': 'Influent_raw', 'Count_num': 'Influent_used'})

    # Effluent samples for plant
    df_eff = df_effluent[df_effluent['Plant'] == plant][['Event', 'Count', 'Count_num']]
    df_eff = df_eff.rename(columns={'Count': 'Effluent_raw', 'Count_num': 'Effluent_used'})

    # Merge on Event to pair samples of same plant and microorganism
    paired = pd.merge(df_inf, df_eff, on='Event', how='inner')

    paired['Plant'] = plant
    paired['Microorganism'] = 'Adenovirus'

    # Exclude pairs where influent_used <= MDL (skip zero/MDL influent)
    paired = paired[paired['Influent_used'] > mdlv]

    # Drop rows with NaN in influent or effluent used counts
    paired = paired.dropna(subset=['Influent_used', 'Effluent_used'])

    # Calculate LRV = log10(influent_used / effluent_used)
    paired['LRV [log10]'] = np.log10(paired['Influent_used'] / paired['Effluent_used'])

    # Determine Effluent detection flag based on raw effluent count
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

# Combine all plants data
df_pairs = pd.concat(rows, ignore_index=True)
df_pairs = df_pairs.sort_values(['Plant', 'Event']).reset_index(drop=True)

# Plot single-panel box plot
plants_order = ['BiSp', 'BrWo', 'NoWe']

# Determine y-axis range (omit outliers)
percentiles = df_pairs['LRV [log10]'].quantile([0.02, 0.98]).values
y_min = math.floor(percentiles[0])
y_max = math.ceil(percentiles[1])

plant_colors = {'BiSp': 'blue', 'BrWo': 'green', 'NoWe': 'orange'}
marker_symbols = {'Non-detect': 'circle-open', 'Detected': 'x'}

fig = make_subplots(rows=1, cols=1, subplot_titles=['Adenovirus'], shared_yaxes=True)

for plant in plants_order:
    df_plant = df_pairs[df_pairs['Plant'] == plant]

    if df_plant.empty:
        continue

    # Box plot without points and no plant legend entry
    fig.add_trace(go.Box(
        y=df_plant['LRV [log10]'],
        x=[plant]*len(df_plant),
        name=plant,
        marker_color=plant_colors.get(plant, 'black'),
        boxmean=True,
        boxpoints=False,
        showlegend=False,
        line=dict(width=2)
    ), row=1, col=1)

    # Scatter for non-detect effluent points (hollow circle)
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
        showlegend=(plant == plants_order[0]),
        legendgroup='Effluent Non-detect'
    ), row=1, col=1)

    # Scatter for detected effluent points (x marker)
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
        showlegend=(plant == plants_order[0]),
        legendgroup='Effluent Detected'
    ), row=1, col=1)

# Update axes formatting
fig.update_yaxes(title_text='LRV [log10]', range=[y_min, y_max], zeroline=True, zerolinewidth=1, zerolinecolor='gray', row=1, col=1)
fig.update_xaxes(title_text='Plant', row=1, col=1)

fig.update_layout(
    font=dict(size=16),
    boxmode='group',
    showlegend=True,
    legend=dict(title='Effluent detection'),
    margin=dict(t=50, b=50, l=50, r=50),
    height=600,
    width=1000,
    template='plotly_white'
)

# Save outputs
base_filename = "lrv_boxplot_adenovirus_by_plant"
html_file = base_filename + ".html"
png_file = base_filename + ".png"
csv_file = base_filename + ".csv"

fig.write_html(html_file, include_plotlyjs='cdn')
fig.write_image(png_file, width=1000, height=600)
df_pairs.to_csv(csv_file, index=False)

print(csv_file)
print(html_file)
print(png_file)
print("lrv_boxplot_adenovirus_by_plant.py")