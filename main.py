#Utviklet av Niklas Brede, Asplan Viak

import time
import arcpy
import pandas as pd
import pathlib
from pathlib import Path

aprx = arcpy.mp.ArcGISProject(r'C:\Users\niklas.brede\OneDrive - Asplan Viak\Documents\Prosjekter\Hydro_skredmal\Skredmal_GIS_NY\Mal_skredfarevurdering.aprx')
project_name = pathlib.Path(aprx.filePath)
directory = str(project_name.parent)

arcpy.env.workspace = r'C:\Users\niklas.brede\OneDrive - Asplan Viak\Documents\Prosjekter\Hydro_skredmal\Skredmal_GIS_NY\Mal_skredfarevurdering.gdb'
arcpy.env.overwriteOutput = True


bkmk = r'Bookmark 1'  # arcpy.GetParameterAsText(0)
excel = r'C:\Users\niklas.brede\OneDrive - Asplan Viak\Documents\Prosjekter\Hydro_skredmal\borehull_mal.xlsx'  # arcpy.GetParameterAsText(1)

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
    # Extract metadata
    koord, antall_prover = extract_coords(sheet_name), extract_nr_tests(sheet_name)
    metadata[sheet_name] = {'Koordinater': koord, 'Antall prøver': antall_prover}

    # Read and process the data
    df = pd.read_excel(excel, sheet_name=sheet_name, skiprows=3)
    df.columns = ['Dybde', 'Materiale', 'Tilstand']
    df = df.iloc[1:].reset_index(drop=True)
    dfs[sheet_name] = df

    # Set local variables
    out_path = r'C:\Users\niklas.brede\OneDrive - Asplan Viak\Documents\Prosjekter\Hydro_skredmal\Skredmal_GIS_NY\output'
    out_name = f"{sheet_name}.shp"
    geometry_type = "POINT"
    spatial_ref = arcpy.SpatialReference(32632)

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
    arcpy.AddField_management(f"{out_path}/{out_name}", 'Dybde', 'DOUBLE')
    arcpy.AddField_management(f"{out_path}/{out_name}", 'Materiale', 'TEXT', field_length=255)
    arcpy.AddField_management(f"{out_path}/{out_name}", 'Tilstand', 'TEXT', field_length=255)
    arcpy.AddField_management(f"{out_path}/{out_name}", 'Str', 'DOUBLE')

    feature_class = f"{out_path}/{out_name}"

    borehole_location_data = [extract_coords(sheet_name)[1], extract_coords(sheet_name)[0], out_name]


    borehole_data = df

    with arcpy.da.InsertCursor(f"{out_path}/{out_name}",
                               ['SHAPE@XY', 'Dybde', 'Materiale', 'Tilstand', 'Str']) as cursor:
        for index, row in df.iterrows():
            xy = (extract_coords(sheet_name)[1], extract_coords(sheet_name)[0])
            size = row['Dybde'] / 100
            cursor.insertRow([xy, row['Dybde'], row['Materiale'], row['Tilstand'], size])
        del cursor















