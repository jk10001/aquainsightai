# filename: detection_prevalence_table_revision.py
import pandas as pd

base_filename = "detection_prevalence_influent_vs_effluent_by_plant_pooled"
csv_filename = base_filename + ".csv"

# Load existing CSV table generated previously
df = pd.read_csv(csv_filename)

# 1) Make sample count columns integers (no decimals)
count_columns = [
    'n_influent_samples [-]',
    'n_influent_detected [-]',
    'n_effluent_samples [-]',
    'n_effluent_detected [-]'
]
for col in count_columns:
    # Convert to integer by rounding first, handle floats gracefully
    df[col] = df[col].round(0).astype(int)

# 2) Round prevalence columns to 1 decimal place
prev_columns = [
    'influent_detection_prevalence [%]',
    'effluent_detection_prevalence [%]'
]
for col in prev_columns:
    df[col] = df[col].round(1)

# 3) Microorganism naming: map existing to preferred capitalization & common names
# Original mapping keys in lowercase, map to proper case/format
microorganism_name_map = {
    'adenovirus': 'Adenovirus',
    'giardia cyst': 'Giardia',
    'crypto oocyst': 'Cryptosporidium',
    'soma coliphage': 'Somatic coliphage',
    'male specific coliphage': 'Male-specific coliphage',
    'total coliform': 'Total coliform',
    'fecal coliform': 'Fecal coliform',
    'e. coli': 'E. coli',
    'aerobic endospores': 'Aerobic endospores'
}

def map_microorganism_name(name):
    name_lower = str(name).strip().lower()
    return microorganism_name_map.get(name_lower, name)

df['Microorganism'] = df['Microorganism'].apply(map_microorganism_name)

# Save the revised CSV to the same filename
df.to_csv(csv_filename, index=False)

print(csv_filename)
print("detection_prevalence_table_revision.py")