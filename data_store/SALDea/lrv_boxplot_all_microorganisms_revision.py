# filename: lrv_boxplot_all_microorganisms_revision.py
import pandas as pd
import numpy as np
import plotly.graph_objs as go
import math

# Load data
file_name = "WWTP_microbial_loads_and_removal.xlsx"
xls = pd.ExcelFile(file_name)
df_data = pd.read_excel(xls, sheet_name="Data")
df_mdl = pd.read_excel(xls, sheet_name="Method detection limits")

# Normalize Plant codes: unify 'BiSP' -> 'BiSp'
df_data['Plant'] = df_data['Plant'].replace({'BiSP': 'BiSp'})

# Microorganisms to include (order matters)
microorganisms_ordered = [
    'Total coliform', 'Fecal coliform', 'E. coli',
    'Somatic coliphage', 'Male-specific coliphage',
    'Aerobic endospores', 'Giardia', 'Cryptosporidium',
    'Adenovirus'
]

# Mapping from IndicatorName to Microorganism output names for included microorganisms
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

# Keep only relevant microorganisms
df_data = df_data[df_data['Microorganism'].isin(microorganisms_ordered)].copy()

# Create lookup for MDL and Units by lower-case organism name
mdl_dict = {row['Organism'].lower(): row['MDL'] for _, row in df_mdl.iterrows()}
units_dict = {row['Organism'].lower(): row['Units'] for _, row in df_mdl.iterrows()}

# Mapping Microorganism to MDL Organism names (for MDL and Units lookup)
microorganism_to_mdl_organism = {
    'Total coliform': 'Total coliform',
    'Fecal coliform': 'Fecal coliform',
    'E. coli': 'E. coli',
    'Somatic coliphage': 'Soma Coliphage',
    'Male-specific coliphage': 'Male Specific Coliphage',
    'Aerobic endospores': 'Aerobic endospores',
    'Giardia': 'Giardia Cyst',
    'Cryptosporidium': 'Crypto Oocyst',
    'Adenovirus': 'Adenovirus',
}

# Add MDL and Units columns
def get_mdl_and_units(microorganism):
    mdl_org = microorganism_to_mdl_organism.get(microorganism)
    if mdl_org:
        mdl_val = mdl_dict.get(mdl_org.lower(), np.nan)
        unit_val = units_dict.get(mdl_org.lower(), '')
        return pd.Series([mdl_val, unit_val])
    return pd.Series([np.nan, ''])

df_data[['MDL', 'Units']] = df_data['Microorganism'].apply(get_mdl_and_units)

# Effluent selection rules:
# Giardia, Cryptosporidium, Somatic coliphage, Male-specific coliphage use 'UF Effluent'
# Others use 'Effluent grab'
uf_effluent_organisms = {'Giardia', 'Cryptosporidium', 'Somatic coliphage', 'Male-specific coliphage'}

# Replace BDL and numeric 0 with organism-specific MDL for Count_num
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

# Separate dataframes by sample type
df_influent = df_data[df_data['SampleType'] == 'Influent grab'].copy()
df_uf_effluent = df_data[df_data['SampleType'] == 'UF Effluent'].copy()
df_grab_effluent = df_data[df_data['SampleType'] == 'Effluent grab'].copy()

# Prepare list of paired data for plotting
rows = []

# Unique plants and microorganisms for pairing
plants = sorted(df_data['Plant'].dropna().unique())

for microorganism in microorganisms_ordered:
    mdl_value = df_data[df_data['Microorganism'] == microorganism]['MDL'].dropna().unique()
    if len(mdl_value) == 0:
        continue
    mdl_value = mdl_value[0]

    for plant in plants:
        # Influent samples with counts
        df_inf = df_influent[(df_influent['Plant'] == plant) & (df_influent['Microorganism'] == microorganism)][['Event', 'Count', 'Count_num']]
        df_inf = df_inf.rename(columns={'Count': 'Influent_raw', 'Count_num': 'Influent_used'})

        # Effluent selection
        if microorganism in uf_effluent_organisms:
            df_eff = df_uf_effluent[(df_uf_effluent['Plant'] == plant) & (df_uf_effluent['Microorganism'] == microorganism)][['Event', 'Count', 'Count_num']]
        else:
            df_eff = df_grab_effluent[(df_grab_effluent['Plant'] == plant) & (df_grab_effluent['Microorganism'] == microorganism)][['Event', 'Count', 'Count_num']]

        df_eff = df_eff.rename(columns={'Count': 'Effluent_raw', 'Count_num': 'Effluent_used'})

        # Merge influent and effluent by Event (and plant, microorganism known)
        paired = pd.merge(df_inf, df_eff, on='Event', how='inner')

        # Add Plant and Microorganism columns
        paired['Plant'] = plant
        paired['Microorganism'] = microorganism

        # Exclude pairs where influent_used <= MDL (ignore zero/MDL influent)
        paired = paired[paired['Influent_used'] > mdl_value]

        # Drop pairs with NaN values in influent or effluent used
        paired = paired.dropna(subset=['Influent_used', 'Effluent_used'])

        # Compute LRV = log10(influent_used / effluent_used)
        paired['LRV [log10]'] = np.log10(paired['Influent_used'] / paired['Effluent_used'])

        # Effluent detection flags based on raw effluent counts
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

# Sort by Microorganism and Event for output neatness
df_pairs = df_pairs.sort_values(['Microorganism', 'Event']).reset_index(drop=True)

# Plot single box plot with all microorganisms, no differentiation by plant
plant_colors = {'BiSp': 'blue', 'BrWo': 'green', 'NoWe': 'orange'}  # Not used for box colors here, but markers can be all black for clarity

marker_symbols = {'Non-detect': 'circle-open', 'Detected': 'x'}

# Determine y-axis range for LRV, excluding outliers
percentiles = df_pairs['LRV [log10]'].quantile([0.02, 0.98]).values
y_min = math.floor(percentiles[0])
y_max = math.ceil(percentiles[1])

import plotly.graph_objs as go

fig = go.Figure()

# Plot box for each microorganism with showlegend=False for all box traces
for micro in microorganisms_ordered:
    df_m = df_pairs[df_pairs['Microorganism'] == micro]
    if df_m.empty:
        continue
    fig.add_trace(go.Box(
        y=df_m['LRV [log10]'],
        name=micro,
        boxmean=True,
        marker_color='blue',
        boxpoints=False,
        showlegend=False  # Hide box trace legend entries
    ))

# Overlay scatter points for all data points with markers based on Effluent detection

# Combine all data points in order of microorganisms (to preserve order on x-axis)
x_categories = microorganisms_ordered
cat_to_x = {cat: i for i, cat in enumerate(x_categories)}

scatter_detect_x = []
scatter_detect_y = []
scatter_nondetect_x = []
scatter_nondetect_y = []

for idx, row in df_pairs.iterrows():
    x_val = row['Microorganism']
    y_val = row['LRV [log10]']
    if row['Effluent_detected_flag'] == 'Non-detect':
        scatter_nondetect_x.append(x_val)
        scatter_nondetect_y.append(y_val)
    elif row['Effluent_detected_flag'] == 'Detected':
        scatter_detect_x.append(x_val)
        scatter_detect_y.append(y_val)

# Add scatter for non-detect
fig.add_trace(go.Scatter(
    x=scatter_nondetect_x,
    y=scatter_nondetect_y,
    mode='markers',
    marker=dict(
        symbol='circle-open',
        color='black',
        size=8,
        line=dict(width=1)
    ),
    name='Effluent Non-detect',
    showlegend=True
))

# Add scatter for detect
fig.add_trace(go.Scatter(
    x=scatter_detect_x,
    y=scatter_detect_y,
    mode='markers',
    marker=dict(
        symbol='x',
        color='black',
        size=8,
        line=dict(width=1)
    ),
    name='Effluent Detected',
    showlegend=True
))

fig.update_layout(
    font=dict(size=16),
    yaxis_title='LRV [log10]',
    xaxis_title='Microorganism',
    xaxis=dict(tickangle=45, categoryorder='array', categoryarray=x_categories),
    boxmode='group',
    margin=dict(t=30, b=100, l=70, r=20),
    template='plotly_white',
    height=600,
    width=1000,
    legend=dict(title='Effluent detection')
)

# Save output files
base_filename = "lrv_boxplot_all_microorganisms"
html_file = base_filename + ".html"
png_file = base_filename + ".png"
csv_file = base_filename + ".csv"

fig.write_html(html_file, include_plotlyjs='cdn')
fig.write_image(png_file, width=1000, height=600)
df_pairs.to_csv(csv_file, index=False)

print(csv_file)
print(html_file)
print(png_file)
print("lrv_boxplot_all_microorganisms_revision.py")