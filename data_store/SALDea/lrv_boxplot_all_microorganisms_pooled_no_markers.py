# filename: lrv_boxplot_all_microorganisms_pooled_no_markers.py
import pandas as pd
import numpy as np
import plotly.graph_objs as go
import math

# Load data from Excel file
file_name = "WWTP_microbial_loads_and_removal.xlsx"
xls = pd.ExcelFile(file_name)
df_data = pd.read_excel(xls, sheet_name="Data")
df_mdl = pd.read_excel(xls, sheet_name="Method detection limits")

# Normalize Plant codes to consistent casing
df_data['Plant'] = df_data['Plant'].replace({'BiSP': 'BiSp'})

# Define full microorganism order for x-axis
microorganism_order = [
    'Total coliform', 'Fecal coliform', 'E. coli',
    'Somatic coliphage', 'Male-specific coliphage',
    'Aerobic endospores', 'Giardia',
    'Cryptosporidium', 'Adenovirus'
]

# Mapping from indicator names (lowercase key fragments) to final microorganism names
indicator_to_microorganism = {
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
    for key, val in indicator_to_microorganism.items():
        if key in ind_name_lower:
            return val
    return None

# Apply mapping
df_data['Microorganism'] = df_data['IndicatorName'].apply(map_indicator_to_microorganism)

# Filter to only include rows with microorganisms in the standard order list (exclude others)
df_data = df_data[df_data['Microorganism'].isin(microorganism_order)].copy()

# Create lookup dictionaries for MDL and Units keyed by lowercase organism names (from Method detection limits sheet)
mdl_dict = {}
units_dict = {}
for _, row in df_mdl.iterrows():
    mdl_dict[row['Organism'].lower()] = row['MDL']
    units_dict[row['Organism'].lower()] = row['Units']

# Map output microorganism names to MDL organism sheet names (to get MDL)
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

# Retrieve MDL for microorganism
def get_mdl(microorganism):
    mdl_org = microorganism_to_mdl_organism.get(microorganism)
    if mdl_org:
        return mdl_dict.get(mdl_org.lower(), np.nan)
    return np.nan

# Add MDL column
df_data['MDL'] = df_data['Microorganism'].apply(get_mdl)

# Define microorganism groups that use UF Effluent as effluent sample (per project rules)
uf_effluent_organisms = {'Giardia', 'Cryptosporidium', 'Somatic coliphage', 'Male-specific coliphage'}

# Replace BDL and numeric zero with organism-specific MDL for calculation convenience
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
    except Exception:
        return np.nan

df_data['Count_num'] = df_data.apply(replace_bdl_zero_with_mdl, axis=1)

# Separate influent and effluent data according to sample type and microorganism
df_influent = df_data[df_data['SampleType'] == 'Influent grab'].copy()
df_uf_effluent = df_data[df_data['SampleType'] == 'UF Effluent'].copy()
df_grab_effluent = df_data[df_data['SampleType'] == 'Effluent grab'].copy()

# Prepare paired LRV data rows for all microorganisms pooled (no plant separation)
paired_rows = []

for microorganism in microorganism_order:
    mdl_value = get_mdl(microorganism)
    if pd.isna(mdl_value):
        continue
    
    # Select influent samples for microorganism (all plants pooled)
    df_inf = df_influent[df_influent['Microorganism'] == microorganism][['Plant', 'Event', 'Count', 'Count_num']]
    df_inf = df_inf.rename(columns={'Count': 'Influent_raw', 'Count_num': 'Influent_used'})

    # Select effluent samples depending on organism group
    if microorganism in uf_effluent_organisms:
        df_eff_sel = df_uf_effluent
    else:
        df_eff_sel = df_grab_effluent
    df_eff = df_eff_sel[df_eff_sel['Microorganism'] == microorganism][['Plant', 'Event', 'Count', 'Count_num']]
    df_eff = df_eff.rename(columns={'Count': 'Effluent_raw', 'Count_num': 'Effluent_used'})

    # Merge influent and effluent samples for same Plant and Event to pair
    paired = pd.merge(df_inf, df_eff, on=['Plant', 'Event'], how='inner')

    # Drop pairs with NaN values in used counts
    paired = paired.dropna(subset=['Influent_used', 'Effluent_used'])

    # Exclude pairs where Influent_used <= MDL (i.e., only use influent > MDL)
    paired = paired[paired['Influent_used'] > mdl_value]

    # Calculate LRV = log10(influent_used / effluent_used)
    paired['LRV [log10]'] = np.log10(paired['Influent_used'] / paired['Effluent_used'])

    # Remove infinite or NaN LRV values if present (shouldn't be normally)
    paired = paired.replace([np.inf, -np.inf], np.nan).dropna(subset=['LRV [log10]'])

    # Add Microorganism column for plotting
    paired['Microorganism'] = microorganism

    # Append results
    paired_rows.append(paired[['Event', 'Plant', 'Microorganism', 'Influent_raw', 'Effluent_raw', 'Influent_used', 'Effluent_used', 'LRV [log10]']])

# Concatenate all microorganism pairs for pooled dataset
df_paired_all = pd.concat(paired_rows, ignore_index=True)

# Prepare box plot colors for each microorganism distinctively (consistent order)
box_colors = {
    'Total coliform': '#1f77b4',
    'Fecal coliform': '#ff7f0e',
    'E. coli': '#2ca02c',
    'Somatic coliphage': '#d62728',
    'Male-specific coliphage': '#9467bd',
    'Aerobic endospores': '#8c564b',
    'Giardia': '#e377c2',
    'Cryptosporidium': '#7f7f7f',
    'Adenovirus': '#bcbd22'
}

# Create box plot with one box per microorganism, pooled across all plants
fig = go.Figure()

# Determine y-axis range based on aggregate 2nd and 98th percentiles of all LRV values (omit extreme outliers)
percentiles_all = df_paired_all['LRV [log10]'].quantile([0.02, 0.98]).values
y_min = math.floor(percentiles_all[0])
y_max = math.ceil(percentiles_all[1])

for microbe in microorganism_order:
    lrv_vals = df_paired_all[df_paired_all['Microorganism'] == microbe]['LRV [log10]']
    if lrv_vals.empty:
        continue
    fig.add_trace(go.Box(
        y=lrv_vals,
        name=microbe,
        boxpoints=False,
        boxmean=True,  # Draw mean as dashed line with standard deviation shading
        line=dict(width=2, color=box_colors.get(microbe, 'black')),
        marker=dict(color=box_colors.get(microbe, 'black')),
        showlegend=False,
        width=0.6  # Added width parameter to make boxes wider per user request
    ))

# Update layout per instructions
fig.update_layout(
    font=dict(size=16),
    yaxis=dict(
        title='LRV [log10]',
        range=[y_min, y_max],
        gridcolor='lightgray',
        zeroline=True,
        zerolinewidth=1,
        zerolinecolor='gray',
        showgrid=True
    ),
    xaxis=dict(
        title='Microorganism',
        showgrid=True,
        gridcolor='lightgray',
        categoryorder='array',
        categoryarray=microorganism_order
    ),
    boxmode='group',
    showlegend=False,
    margin=dict(t=30, b=60, l=60, r=40),
    template='plotly_white',
    height=600,
    width=1000,
)

# File base name for outputs
base_filename = "lrv_boxplot_all_microorganisms_pooled_no_markers"
html_file = base_filename + ".html"
png_file = base_filename + ".png"
csv_file = base_filename + ".csv"
py_file = base_filename + ".py"

# Save outputs
fig.write_html(html_file, include_plotlyjs='cdn')
fig.write_image(png_file, width=1000, height=600)
df_paired_all.to_csv(csv_file, index=False)

print(csv_file)
print(html_file)
print(png_file)
print(py_file)