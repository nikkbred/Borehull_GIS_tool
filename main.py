import arcpy
import pandas as pd
import pathlib
import os

aprx = arcpy.mp.ArcGISProject(r"C:\Users\niklas.brede\OneDrive - Asplan Viak\Documents\Prosjekter\Hydro_skredmal\Miljøgeologi-GISmal\Mal_skredfarevurdering.aprx")
project_name = pathlib.Path(aprx.filePath)
directory = str(project_name.parent)
arcpy.env.workspace = directory
arcpy.env.overwriteOutput = True
excel = r"C:\Users\niklas.brede\OneDrive - Asplan Viak\Documents\Prosjekter\Hydro_skredmal\grunnforurensing.xlsx"
xls = pd.ExcelFile(excel)


def get_intervention_class(pollutant, value):
    if pollutant == 'As':
        if value < 8:
            return "Meget god"
        elif 8 <= value < 20:
            return "God"
        elif 20 <= value < 50:
            return "Moderat"
        elif 50 <= value < 600:
            return "Dårlig"
        elif 600 <= value <= 1000:
            return "Svært dårlig"

    elif pollutant == 'Pb':
        if value < 60:
            return "Meget god"
        elif 60 <= value < 100:
            return "God"
        elif 100 <= value < 300:
            return "Moderat"
        elif 300 <= value < 700:
            return "Dårlig"
        elif 700 <= value <= 2500:
            return "Svært dårlig"

    elif pollutant == 'Cd':
        if value < 1.5:
            return "Meget god"
        elif 1.5 <= value < 10:
            return "God"
        elif 10 <= value < 15:
            return "Moderat"
        elif 15 <= value < 30:
            return "Dårlig"
        elif 30 <= value <= 1000:
            return "Svært dårlig"

    elif pollutant == 'Hg':
        if value < 1:
            return "Meget god"
        elif 1 <= value < 2:
            return "God"
        elif 2 <= value < 4:
            return "Moderat"
        elif 4 <= value < 10:
            return "Dårlig"
        elif 10 <= value <= 1000:
            return "Svært dårlig"

    elif pollutant == 'Cu':
        if value < 100:
            return "Meget god"
        elif 100 <= value < 200:
            return "God"
        elif 200 <= value < 1000:
            return "Moderat"
        elif 1000 <= value < 8500:
            return "Dårlig"
        elif 8500 <= value <= 25000:
            return "Svært dårlig"

    elif pollutant == 'Zn':
        if value < 200:
            return "Meget god"
        elif 200 <= value < 500:
            return "God"
        elif 500 <= value < 1000:
            return "Moderat"
        elif 1000 <= value < 5000:
            return "Dårlig"
        elif 5000 <= value <= 25000:
            return "Svært dårlig"

    elif pollutant == 'Cr':
        if value < 50:
            return "Meget god"
        elif 50 <= value < 200:
            return "God"
        elif 200 <= value < 500:
            return "Moderat"
        elif 500 <= value < 2800:
            return "Dårlig"
        elif 2800 <= value <= 25000:
            return "Svært dårlig"

    elif pollutant == 'Ni':
        if value < 60:
            return "Meget god"
        elif 60 <= value < 135:
            return "God"
        elif 135 <= value < 200:
            return "Moderat"
        elif 200 <= value < 1200:
            return "Dårlig"
        elif 1200 <= value <= 2500:
            return "Svært dårlig"

    elif pollutant == 'PCB7':
        if value < 0.01:
            return "Meget god"
        elif 0.01 <= value < 0.5:
            return "God"
        elif 0.5 <= value < 1:
            return "Moderat"
        elif 1 <= value < 5:
            return "Dårlig"
        elif 5 <= value <= 50:
            return "Svært dårlig"

    elif pollutant == 'DDT':
        if value < 0.04:
            return "Meget god"
        elif 0.04 <= value < 4:
            return "God"
        elif 4 <= value < 12:
            return "Moderat"
        elif 12 <= value < 30:
            return "Dårlig"
        elif 30 <= value <= 50:
            return "Svært dårlig"

    elif pollutant == 'PAH-16EPA':
        if value < 2:
            return "Meget god"
        elif 2 <= value < 8:
            return "God"
        elif 8 <= value < 50:
            return "Moderat"
        elif 50 <= value < 150:
            return "Dårlig"
        elif 150 <= value <= 2500:
            return "Svært dårlig"

    elif pollutant == 'BAP':
        if value < 0.1:
            return "Meget god"
        elif 0.1 <= value < 0.5:
            return "God"
        elif 0.5 <= value < 5:
            return "Moderat"
        elif 5 <= value < 15:
            return "Dårlig"
        elif 15 <= value <= 100:
            return "Svært dårlig"

    elif pollutant == 'AlC8_10':
        if value <= 10:
            return "Meget god"
        elif 10 < value <= 40:
            return "God"
        elif 40 < value <= 50:
            return "Moderat"
        elif 50 < value <= 20000:
            return "Dårlig"

    elif pollutant == 'AlC10_12':
        if value < 50:
            return "Meget god"
        elif 50 <= value < 60:
            return "God"
        elif 60 <= value < 130:
            return "Moderat"
        elif 130 <= value < 300:
            return "Dårlig"
        elif 300 <= value <= 20000:
            return "Svært dårlig"

    elif pollutant == 'AlC12_16' and 'AlC16_35':
        if value < 100:
            return "Meget god"
        elif 100 <= value < 300:
            return "God"
        elif 300 <= value < 600:
            return "Moderat"
        elif 600 <= value < 2000:
            return "Dårlig"
        elif 2000 <= value <= 20000:
            return "Svært dårlig"

    elif pollutant == 'DEHP':
        if value < 2.8:
            return "Meget god"
        elif 2.8 <= value < 25:
            return "God"
        elif 25 <= value < 40:
            return "Moderat"
        elif 40 <= value < 60:
            return "Dårlig"
        elif 60 <= value <= 5000:
            return "Svært dårlig"

    elif pollutant == 'Dioksiner/furaner':
        if value < 0.00001:
            return "Meget god"
        elif 0.00001 <= value < 0.00002:
            return "God"
        elif 0.00002 <= value < 0.0001:
            return "Moderat"
        elif 0.0001 <= value < 0.00036:
            return "Dårlig"
        elif 0.00036 <= value <= 0.015:
            return "Svært dårlig"

    elif pollutant == 'Fenol':
        if value < 0.1:
            return "Meget god"
        elif 0.1 <= value < 4:
            return "God"
        elif 4 <= value < 40:
            return "Moderat"
        elif 40 <= value < 400:
            return "Dårlig"
        elif 400 <= value <= 25000:
            return "Svært dårlig"

    elif pollutant == 'BENZ':
        if value < 0.01:
            return "Meget god"
        elif 0.01 <= value < 0.015:
            return "God"
        elif 0.015 <= value < 0.04:
            return "Moderat"
        elif 0.04 <= value < 0.05:
            return "Dårlig"
        elif 0.05 <= value <= 1000:
            return "Svært dårlig"
    else:
        return "Pollutant not found"




# Dictionary to store the DataFrames and metadata
dfs = {}

sheet_name = 'provepunkter'

# Read and process the data
df = pd.read_excel(excel, sheet_name=sheet_name)

df = df.replace(',', '.', regex=True)
df['Koordinat_O'] = pd.to_numeric(df['Koordinat_O'], errors='coerce')
df['Koordinat_N'] = pd.to_numeric(df['Koordinat_N'], errors='coerce')

dfs[sheet_name] = df

coords = df[['Koordinat_O','Koordinat_N']].drop_duplicates()


# Initialize a dictionary to hold the DataFrames for each unique coordinate
dfs_by_coord = {}

# Iterate over unique coordinates
for coord in coords.itertuples(index=False, name=None):
    # Filter the original DataFrame for the current coordinate
    df_filtered = df[(df['Koordinat_O'] == coord[0]) & (df['Koordinat_N'] == coord[1])]

    # Add the filtered DataFrame to the dictionary with the coordinate as the key
    dfs_by_coord[coord] = df_filtered

# Now dfs_by_coord contains DataFrames for each unique coordinate
for coord, df_part in dfs_by_coord.items():
    print(f"Coordinate: {coord}")
    print(df_part)
    print()

gdb_name = "Mal_skredfarevurdering.gdb"

# Define the path to the new geodatabase
gdb_path = f"{directory}/{gdb_name}"

# Loop through each coordinate and dataframe pair
for coord, df in dfs_by_coord.items():
    point_name = df['Provepunkt_Navn'].iloc[0]
    feature_class_name = f"{point_name}"
    feature_class_path = f"{gdb_path}/{feature_class_name}"

    # Create a new feature class for each point
    arcpy.CreateFeatureclass_management(gdb_path, feature_class_name, "POINT",
                                        spatial_reference=arcpy.SpatialReference(25833))

    # Add fields for the attributes
    arcpy.AddField_management(feature_class_path, "Dybde", "SHORT")
    arcpy.AddField_management(feature_class_path, "Tiltaksklasse", "SHORT")
    arcpy.AddField_management(feature_class_path, "Dim_stoff", "TEXT")


    with arcpy.da.InsertCursor(feature_class_path, ["SHAPE@XY", "Dybde", "Tiltaksklasse", "Dim_stoff"]) as cursor:
        for idx, row in df.iterrows():
            dybde_n = row['Dyp_nedre']
            tiltk = row['WKID']
            stoff = row['Stoff_ID']
            cursor.insertRow([(coord[0], coord[1]), dybde_n, tiltk, stoff])



#TODO: Legg til to kolloner i df; tiltaksklasse og hvilket stoff som er dimensjonerende for denne tiltaksklassen.