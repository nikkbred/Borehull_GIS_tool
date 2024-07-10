import arcpy
import pandas as pd
import pathlib
import os
import time

aprx = arcpy.mp.ArcGISProject('CURRENT')
project_name = pathlib.Path(aprx.filePath)
directory = str(project_name.parent)
arcpy.env.workspace = directory
arcpy.env.overwriteOutput = True
excel = arcpy.GetParameterAsText(0)
xls = pd.ExcelFile(excel)


# Function to extract metadata from a sheet
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


def get_intervention_class_trondelag(pollutant, value):
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
    elif pollutant == 'Ni':
        if value < 75:
            return 1
        elif 75 <= value < 135:
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
if arcpy.GetParameter(1) == True:
    df['Tilstandsklasse'] = df.apply(lambda row: get_intervention_class_trondelag(row['Stoff_ID'], row['Verdi']),
                                     axis=1)
else:
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
for i in range(len(aprx.listMaps())):
    map_obj = aprx.listMaps()[i].name
    print(str(i) + ' ' + map_obj)
gdb_name = "Miljøgeologi_GISmal.gdb"
# Define the path to the new geodatabase
gdb_path = f"{directory}/Provepunkt"
koord_id = 0
# Loop through each coordinate and dataframe pair
for coord, df in dfs_by_coord.items():
    koord_id += 1
    point_name = f"Provepunkt_{koord_id}"
    feature_class_name = point_name
    feature_class_path = f"{gdb_path}/{feature_class_name}"
    SRtxt = arcpy.GetParameterAsText(1)
    SR = arcpy.SpatialReference()

    # Create a new feature class for each point
    arcpy.CreateFeatureclass_management(gdb_path, feature_class_name, "POINT",
                                        spatial_reference=SR.loadFromString(SRtxt))
    # Add fields for the attributes
    arcpy.AddField_management(feature_class_path, "Koord_ID", "SHORT")
    arcpy.AddField_management(feature_class_path, "Provepunkt", "TEXT")
    arcpy.AddField_management(feature_class_path, "Dybde", "SHORT")
    arcpy.AddField_management(feature_class_path, "Tilstandkl", "SHORT")
    arcpy.AddField_management(feature_class_path, "Dim_stoff", "TEXT")
    with arcpy.da.InsertCursor(feature_class_path,
                               ["SHAPE@XY", "Koord_ID", "Provepunkt", "Dybde", "Tilstandkl", "Dim_stoff"]) as cursor:
        for idx, row in df.iterrows():
            dybde_n = row['Dyp_nedre']
            provenavn = row['Provepunkt_Navn']
            tilstdkl = row['H_tilstand']
            stoff = row['Dim_stoff']
            cursor.insertRow([(coord[0], coord[1]), koord_id, provenavn, dybde_n, tilstdkl, stoff])
        del cursor
time.sleep(3)


def _set_symbology(path_lyrx, _map):
    symbology_layer = path_lyrx
    new_lyr_file = arcpy.mp.LayerFile(symbology_layer)
    new_lyr = new_lyr_file.listLayers()[0]
    old_lyr = _map.listLayers()[0]
    # self._unselect_lyr(_map)
    old_lyr_name = old_lyr.name
    new_lyr.updateConnectionProperties(new_lyr.connectionProperties, old_lyr.connectionProperties)
    new_lyr.name = old_lyr_name
    new_lyr_file.save()
    _map.insertLayer(old_lyr, new_lyr_file)
    _map.removeLayer(old_lyr)


# Adding data to the map
map_obj = aprx.listMaps()[4]
map_obj2 = aprx.listMaps()[3]
for i in range(koord_id):
    feature_class_path = os.path.join(directory, f"Provepunkt/Provepunkt_{i + 1}.shp")
    try:
        tilst_map = map_obj2.addDataFromPath(feature_class_path)
        layer = map_obj.addDataFromPath(feature_class_path)
        _set_symbology("LYR\Tilstandsklasser.lyrx", map_obj2)


    except RuntimeError as e:
        arcpy.AddMessage(f"Failed to add data for Provepunkt_{i}: {e}")

try:
    for i in range(len(aprx.listLayouts())):
        if aprx.listLayouts()[i].name == 'Oversiktskart':
            layout = aprx.listLayouts()[i]
        else:
            pass
except:
    arcpy.AddMessage('Layouten Oversiktskart ble ikke funnet')

legend = layout.listElements('LEGEND_ELEMENT', 'Tegnforklaring')[0]

seen = set()
items_to_remove = []

for itm in legend.items:
    if itm.name in seen:
        items_to_remove.append(itm)
    else:
        seen.add(itm.name)

for itm in items_to_remove:
    legend.removeItem(itm)

lyt_cim = layout.getDefinition('V2')

for elm in lyt_cim.elements:
    if elm.name == "Tegnforklaring":
        for itm in elm.items:
            itm.showLayerName = False

layout.setDefinition(lyt_cim)

aprx.save()