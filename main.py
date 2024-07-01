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
            return 1
        elif 8 <= value < 20:
            return 2
        elif 20 <= value < 50:
            return 3
        elif 50 <= value < 600:
            return 4
        elif 600 <= value <= 1000:
            return 5

    elif pollutant == 'Pb':
        if value < 60:
            return 1
        elif 60 <= value < 100:
            return 2
        elif 100 <= value < 300:
            return 3
        elif 300 <= value < 700:
            return 4
        elif 700 <= value <= 2500:
            return 5

    elif pollutant == 'Cd':
        if value < 1.5:
            return 1
        elif 1.5 <= value < 10:
            return 2
        elif 10 <= value < 15:
            return 3
        elif 15 <= value < 30:
            return 4
        elif 30 <= value <= 1000:
            return 5

    elif pollutant == 'Hg':
        if value < 1:
            return 1
        elif 1 <= value < 2:
            return 2
        elif 2 <= value < 4:
            return 3
        elif 4 <= value < 10:
            return 4
        elif 10 <= value <= 1000:
            return 5

    elif pollutant == 'Cu':
        if value < 100:
            return 1
        elif 100 <= value < 200:
            return 2
        elif 200 <= value < 1000:
            return 3
        elif 1000 <= value < 8500:
            return 4
        elif 8500 <= value <= 25000:
            return 5

    elif pollutant == 'Zn':
        if value < 200:
            return 1
        elif 200 <= value < 500:
            return 2
        elif 500 <= value < 1000:
            return 3
        elif 1000 <= value < 5000:
            return 4
        elif 5000 <= value <= 25000:
            return 5

    elif pollutant == 'Cr':
        if value < 50:
            return 1
        elif 50 <= value < 200:
            return 2
        elif 200 <= value < 500:
            return 3
        elif 500 <= value < 2800:
            return 4
        elif 2800 <= value <= 25000:
            return 5

    elif pollutant == 'Ni':
        if value < 60:
            return 1
        elif 60 <= value < 135:
            return 2
        elif 135 <= value < 200:
            return 3
        elif 200 <= value < 1200:
            return 4
        elif 1200 <= value <= 2500:
            return 5

    elif pollutant == 'PCB7':
        if value < 0.01:
            return 1
        elif 0.01 <= value < 0.5:
            return 2
        elif 0.5 <= value < 1:
            return 3
        elif 1 <= value < 5:
            return 4
        elif 5 <= value <= 50:
            return 5

    elif pollutant == 'DDT':
        if value < 0.04:
            return 1
        elif 0.04 <= value < 4:
            return 2
        elif 4 <= value < 12:
            return 3
        elif 12 <= value < 30:
            return 4
        elif 30 <= value <= 50:
            return 5

    elif pollutant == 'PAH-16EPA':
        if value < 2:
            return 1
        elif 2 <= value < 8:
            return 2
        elif 8 <= value < 50:
            return 3
        elif 50 <= value < 150:
            return 4
        elif 150 <= value <= 2500:
            return 5

    elif pollutant == 'BAP':
        if value < 0.1:
            return 1
        elif 0.1 <= value < 0.5:
            return 2
        elif 0.5 <= value < 5:
            return 3
        elif 5 <= value < 15:
            return 4
        elif 15 <= value <= 100:
            return 5

    elif pollutant == 'AlC8_10':
        if value <= 10:
            return 1
        elif 10 < value <= 40:
            return 2
        elif 40 < value <= 50:
            return 3
        elif 50 < value <= 20000:
            return 4

    elif pollutant == 'AlC10_12':
        if value < 50:
            return 1
        elif 50 <= value < 60:
            return 2
        elif 60 <= value < 130:
            return 3
        elif 130 <= value < 300:
            return 4
        elif 300 <= value <= 20000:
            return 5

    elif pollutant == 'AlC12_16' and 'AlC16_35':
        if value < 100:
            return 1
        elif 100 <= value < 300:
            return 2
        elif 300 <= value < 600:
            return 3
        elif 600 <= value < 2000:
            return 4
        elif 2000 <= value <= 20000:
            return 5

    elif pollutant == 'DEHP':
        if value < 2.8:
            return 1
        elif 2.8 <= value < 25:
            return 2
        elif 25 <= value < 40:
            return 3
        elif 40 <= value < 60:
            return 4
        elif 60 <= value <= 5000:
            return 5

    elif pollutant == 'Dioksiner/furaner':
        if value < 0.00001:
            return 1
        elif 0.00001 <= value < 0.00002:
            return 2
        elif 0.00002 <= value < 0.0001:
            return 3
        elif 0.0001 <= value < 0.00036:
            return 4
        elif 0.00036 <= value <= 0.015:
            return 5

    elif pollutant == 'Fenol':
        if value < 0.1:
            return 1
        elif 0.1 <= value < 4:
            return 2
        elif 4 <= value < 40:
            return 3
        elif 40 <= value < 400:
            return 4
        elif 400 <= value <= 25000:
            return 5

    elif pollutant == 'BENZ':
        if value < 0.01:
            return 1
        elif 0.01 <= value < 0.015:
            return 2
        elif 0.015 <= value < 0.04:
            return 3
        elif 0.04 <= value < 0.05:
            return 4
        elif 0.05 <= value <= 1000:
            return 5
    else:
        return 0



# Dictionary to store the DataFrames and metadata
dfs = {}

sheet_name = 'provepunkter'

# Read and process the data
df = pd.read_excel(excel, sheet_name=sheet_name)

df = df.replace(',', '.', regex=True)
df['Koordinat_O'] = pd.to_numeric(df['Koordinat_O'], errors='coerce')
df['Koordinat_N'] = pd.to_numeric(df['Koordinat_N'], errors='coerce')
df['Verdi'] = pd.to_numeric(df['Verdi'], errors='coerce')

dfs[sheet_name] = df

# Adds column to DataFrame containing the contamination class and applicable contamination type

df['Tilstandsklasse'] = df.apply(lambda row: get_intervention_class(row['Stoff_ID'], row['Verdi']), axis=1)


coords = df[['Koordinat_O', 'Koordinat_N']].drop_duplicates()

# Initialize a dictionary to hold the DataFrames for each unique coordinate
dfs_by_coord = {}

# Iterate over unique coordinates
for coord in coords.itertuples(index=False, name=None):
    # Filter the original DataFrame for the current coordinate
    df_filtered = df[(df['Koordinat_O'] == coord[0]) & (df['Koordinat_N'] == coord[1])]


    # Calculate the highest pollution class for each depth
    def highest_pollution_class(group):
        highest_class = group['Tilstandsklasse'].max()
        substances = group[group['Tilstandsklasse'] == highest_class]['Stoff_ID'].tolist()
        return pd.Series({
            'H_tilstand': highest_class,
            'Dim_stoff': ', '.join(substances)
        })


    highest_pollution = df_filtered.groupby('Dyp_nedre').apply(highest_pollution_class).reset_index()
    df_filtered = df_filtered.merge(highest_pollution, on='Dyp_nedre', how='left')

    # Add the filtered DataFrame to the dictionary with the coordinate as the key
    dfs_by_coord[coord] = df_filtered

# Example of how to access and print the processed DataFrame for a specific coordinate
for coord, df in dfs_by_coord.items():
    print(f"Coordinate: {coord}")
    print(df)



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
    arcpy.AddField_management(feature_class_path, "Tilstand", "SHORT")
    arcpy.AddField_management(feature_class_path, "Dim_stoff", "TEXT")


    with arcpy.da.InsertCursor(feature_class_path, ["SHAPE@XY", "Dybde", "Tilstand", "Dim_stoff"]) as cursor:
        for idx, row in df.iterrows():
            dybde_n = row['Dyp_nedre']
            tiltk = row['H_tilstand']
            stoff = row['Dim_stoff']
            cursor.insertRow([(coord[0], coord[1]), dybde_n, tiltk, stoff])



#TODO: Legg til to kolloner i df; tiltaksklasse og hvilket stoff som er dimensjonerende for denne tiltaksklassen. + legg til  kart(ene)