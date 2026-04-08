# filename: population_projection_20yr.py
import pandas as pd
import plotly.graph_objects as go

# Load population data
file_name = "WWTP_medium.xlsx"
pop_sheet = "population"
df_pop = pd.read_excel(file_name, sheet_name=pop_sheet)

# Restrict to 2014-2017 historical years
df_hist = df_pop[(df_pop["Year"] >= 2014) & (df_pop["Year"] <= 2017)].copy()

# Calculate CAGR over 2014-2017
P0 = df_hist.loc[df_hist["Year"] == 2014, "Serviced Population"].values[0]
Pn = df_hist.loc[df_hist["Year"] == 2017, "Serviced Population"].values[0]
n = 2017 - 2014  # =3
cagr_fraction = (Pn / P0) ** (1 / n) - 1
cagr_percent = cagr_fraction * 100

# Projection years 2018-2037 (20 years)
proj_years = list(range(2018, 2038))
pop_2017 = Pn
proj_pops = [pop_2017 * (1 + cagr_fraction) ** (y - 2017) for y in proj_years]

# Compose DataFrame for projection
df_proj = pd.DataFrame({
    "Year": proj_years,
    "Serviced Population": proj_pops
})

# Assign Series labels
df_hist["Series"] = "Historical"
df_proj["Series"] = "Projected"

# Add growth rate column to both
df_hist["Applied_annual_growth_rate_percent"] = cagr_percent
df_proj["Applied_annual_growth_rate_percent"] = cagr_percent

# Combine historical and projected
df_combined = pd.concat([df_hist, df_proj], ignore_index=True)

# Save CSV of combined data
csv_filename = "population_projection_20yr.csv"
df_combined.to_csv(csv_filename, index=False)

# Create the plotly line chart
fig = go.Figure()

# Historical line (solid)
df_hist_sorted = df_hist.sort_values("Year")
fig.add_trace(go.Scatter(
    x=df_hist_sorted["Year"],
    y=df_hist_sorted["Serviced Population"],
    mode="lines+markers",
    name="Historical",
    line=dict(color="blue", dash="solid"),
    marker=dict(symbol="circle", size=8)
))

# Projected line (dashed)
df_proj_sorted = df_proj.sort_values("Year")
fig.add_trace(go.Scatter(
    x=df_proj_sorted["Year"],
    y=df_proj_sorted["Serviced Population"],
    mode="lines+markers",
    name="Projected",
    line=dict(color="orange", dash="dash"),
    marker=dict(symbol="circle", size=8)
))

# Update layout
fig.update_layout(
    xaxis_title="Year",
    yaxis_title="Serviced population",
    font=dict(size=16),
    template="plotly_white",
    yaxis=dict(showgrid=True, zeroline=False),
    legend=dict(
        bordercolor="black",
        borderwidth=1
    )
)

# Save outputs
base_filename = "population_projection_20yr"
html_filename = base_filename + ".html"
png_filename = base_filename + ".png"

fig.write_html(html_filename, include_plotlyjs='cdn', full_html=True)
fig.write_image(png_filename, width=1000, height=600)

# Print output filenames
print(base_filename + ".py")
print(csv_filename)
print(html_filename)
print(png_filename)