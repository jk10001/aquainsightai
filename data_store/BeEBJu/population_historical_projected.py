# filename: population_historical_projected.py
import pandas as pd
import numpy as np
import plotly.graph_objects as go

# Load population data
file_name = "WWTP_medium.xlsx"
sheet_name = "population"
df_pop = pd.read_excel(file_name, sheet_name=sheet_name)

# Restrict to historical years 2014-2017
df_hist = df_pop[(df_pop["Year"] >= 2014) & (df_pop["Year"] <= 2017)].copy()

# Fit linear regression manually (no sklearn)
# Define X and y
X_hist = df_hist["Year"].values
y_hist = df_hist["Serviced Population"].values

# Calculate slope and intercept for y = a * X + b
n = len(X_hist)
sum_x = X_hist.sum()
sum_y = y_hist.sum()
sum_xx = (X_hist**2).sum()
sum_xy = (X_hist * y_hist).sum()

denominator = n*sum_xx - sum_x**2
if denominator == 0:
    a = 0
    b = y_hist.mean()
else:
    a = (n*sum_xy - sum_x*sum_y) / denominator
    b = (sum_y*sum_xx - sum_x*sum_xy) / denominator

# Years for projection 2014 to 2037
years_proj = np.arange(2014, 2038)

# Predict population
pop_proj = a * years_proj + b
pop_proj_rounded = np.round(pop_proj).astype(int)

# Create DataFrame for output data
# For historical years, show actual population
df_hist_indexed = df_hist.set_index("Year")
population_hist_values = []
for year in years_proj:
    if year in df_hist_indexed.index:
        population_hist_values.append(df_hist_indexed.loc[year, "Serviced Population"])
    else:
        population_hist_values.append(np.nan)

df_out = pd.DataFrame({
    "Year": years_proj,
    "Population_historical": population_hist_values,
    "Population_projected": pop_proj_rounded
})

# Prepare Plotly figure
fig = go.Figure()

# Historical population line and markers
fig.add_trace(go.Scatter(
    x=df_out["Year"],
    y=df_out["Population_historical"],
    mode="lines+markers",
    name="Historical population",
    line=dict(color="blue", width=2, dash="solid"),
    marker=dict(symbol="circle", size=8)
))

# Projected population line and markers
fig.add_trace(go.Scatter(
    x=df_out["Year"],
    y=df_out["Population_projected"],
    mode="lines+markers",
    name="Projected population",
    line=dict(color="orange", width=2, dash="dash"),
    marker=dict(symbol="circle-open", size=8)
))

fig.update_layout(
    xaxis_title="Year",
    yaxis_title="Serviced population (persons)",
    yaxis=dict(showgrid=True, zeroline=False),
    font=dict(size=16),
    template="plotly_white",
    legend=dict(
        bordercolor="black",
        borderwidth=1
    )
)

# Save outputs
base_filename = "population_historical_projected"
csv_filename = base_filename + ".csv"
html_filename = base_filename + ".html"
png_filename = base_filename + ".png"

df_out.to_csv(csv_filename, index=False)
fig.write_html(html_filename, include_plotlyjs='cdn', full_html=True)
fig.write_image(png_filename, width=1000, height=600)

# Print filenames to terminal
print(base_filename + ".py")
print(csv_filename)
print(html_filename)
print(png_filename)