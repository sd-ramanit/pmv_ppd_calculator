import pandas as pd
from pythermalcomfort.models import pmv_ppd
from pythermalcomfort.utilities import v_relative, clo_dynamic

# Load your dataset
file_path = "/Users/manit/Smart-Environmental-and-Energy-Saving-Technologies-in-Building/dataset.xlsx" # Example
df = pd.read_excel(file_path)

# Define constants
met = 1.0  # Bedroom
clo = 0.5  # Light indoor clothing
v = 0.1    # Assumed airspeed in m/s

# Function to calculate PMV and PPD for each row
def calculate_thermal_comfort(row):
    tdb = row['Temperature (°C)']
    rh = row['RH (%)']
    tr = tdb  # Assuming mean radiant temperature equals dry-bulb temperature
    v_r = v_relative(v=v, met=met)
    clo_d = clo_dynamic(clo=clo, met=met)
    results = pmv_ppd(tdb=tdb, tr=tr, vr=v_r, rh=rh, met=met, clo=clo_d)
    return results

# Apply the function to your dataset
df['PMV_PPD'] = df.apply(calculate_thermal_comfort, axis=1)

# Extract PMV and PPD into separate columns
df['PMV'] = df['PMV_PPD'].apply(lambda x: x['pmv'])
df['PPD'] = df['PMV_PPD'].apply(lambda x: x['ppd'])

# Save the results to a new Excel file
df.to_excel("thermal_comfort_analysis.xlsx", index=False)

print("Thermal comfort analysis completed. Results saved to 'thermal_comfort_analysis.xlsx'.")
