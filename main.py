import arcpy
import pandas as pd
import pathlib
import os

aprx = arcpy.mp.ArcGISProject(r"C:\Users\niklas.brede\OneDrive - Asplan Viak\Documents\Prosjekter\Hydro_skredmal\Miljøgeologi-GISmal\Mal_skredfarevurdering.aprx")
project_name = pathlib.Path(aprx.filePath)
directory = str(project_name.parent)
arcpy.env.workspace = directory
arcpy.env.overwriteOutput = True
excel = 'path'
xls = pd.ExcelFile(excel)


# Function to extract metadata from a sheet
def extract_nr_tests(sheet_name):
    df = pd.read_excel(excel, sheet_name=sheet_name, nrows=3, header=None)
    antall_prover = df.iat[2, 1]
    return antall_prover


def extract_coords(sheet_name):
    df = pd.read_excel(excel, sheet_name=sheet_name, nrows=3, header=None)
    koord = (df.iat[1, 1], df.iat[1, 2])
    return koord


# Dictionary to store the DataFrames and metadata
dfs = {}
metadata = {}
# Loop through each sheet and create a DataFrame
for sheet_name in xls.sheet_names:
    global count
    count = len(xls.sheet_names)

    # Extract metadata
    koord, antall_prover = extract_coords(sheet_name), extract_nr_tests(sheet_name)
    metadata[sheet_name] = {'Koordinater': koord, 'Antall prøver': antall_prover}

    # Read and process the data
    df = pd.read_excel(excel, sheet_name=sheet_name, skiprows=3)
    df = df.dropna(axis=1, how='all')
    df.columns = df.iloc[0]
    df = df[1:].reset_index(drop=True)
    dfs[sheet_name] = df
    arcpy.AddMessage(project_name)
    arcpy.AddMessage(df)

    # Set local variables
    out_path = os.path.join(directory, 'Borehull')
    out_name = f"{sheet_name}.shp"

    geometry_type = "POINT"
    spatial_ref = arcpy.SpatialReference(32633)
    template = None
    has_m = "DISABLED"
    has_z = "ENABLED"
    config_keyword = None
    spatial_grid_1 = None
    spatial_grid_2 = None
    spatial_grid_3 = None
    out_alias = None
    oid_type = None
    arcpy.management.CreateFeatureclass(out_path, out_name, geometry_type, template, has_m, has_z, spatial_ref,
                                        config_keyword, spatial_grid_1, spatial_grid_2, spatial_grid_3, out_alias,
                                        oid_type)
    # Add fields to the feature class
    for col in df.columns:
        if col.strip():  # Ensure the column name is not empty
            if df[col].dtype == 'object':
                field_type = 'TEXT'
                arcpy.AddField_management(f"{out_path}/{out_name}", col, field_type, field_length=255)
            elif df[col].dtype in ['int64', 'float64']:
                field_type = 'DOUBLE'
                arcpy.AddField_management(f"{out_path}/{out_name}", col, field_type)
    feature_class = f"{out_path}/{out_name}"
    borehole_location_data = [extract_coords(sheet_name)[1], extract_coords(sheet_name)[0], out_name]
    borehole_data = df
    with arcpy.da.InsertCursor(f"{out_path}/{out_name}", ['SHAPE@XY'] + list(df.columns)) as cursor:
        for index, row in df.iterrows():
            xy = (extract_coords(sheet_name)[1], extract_coords(sheet_name)[0])
            cursor.insertRow([xy] + list(row))
        del cursor


#Legges inn i hvert kart
# for j in range(len(aprx.listMaps())):
#   for i in range(count):
#      map_obj = aprx.listMaps()[j]
#     layer = map_obj.addDataFromPath(os.path.join(directory, f"Borehull\\BH{i+1}.shp"))
for i in range(count):
    map_obj = aprx.listMaps()[4]
    layer = map_obj.addDataFromPath(os.path.join(directory, f"Borehull\\BH{i + 1}.shp"))
