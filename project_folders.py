# -*- coding: utf-8 -*-
"""
Project folder setup script for Kevin McManigal and the Tigers Forever
mapping project
Author: Garin Wally
Date: March 7, 2015 (rewritten/heavily revised March 2018)
License: MIT
ArcGIS and related products are property of ESRI: all rights reserved
"""


# =============================================================================
# Imports
# =============================================================================

import datetime as dt
import logging
import os
import shutil
import sys
from ConfigParser import SafeConfigParser
from subprocess import Popen, PIPE
from time import sleep

import win32com.client as win32


# =============================================================================
# Setup Log File
# =============================================================================

# Output log file & message format
if not os.path.exists("logs"):
    os.mkdir("logs")
log_name = dt.datetime.now().strftime("logs/log-%Y-%m-%d_%H.%M.txt")

logging.basicConfig(filename=log_name,
                    filemode="w",
                    format="[%(asctime)s] %(levelname)-8s : %(message)s",
                    level=logging.DEBUG)


# =============================================================================
# Configuration
# =============================================================================

logging.info("Configuring...")

# Open configs
try:
    # Read main config file config.ini
    CONFIG = SafeConfigParser()
    CONFIG.read("config/config.ini")

    # Get datasets from dataset.ini config file
    DATASETS = SafeConfigParser()
    DATASETS.read("config/datasets.ini")
except Exception as err:
    logging.error(err)
    raise err

# =============================================================================
# Get config options

# Set mode
MODE = CONFIG.get("DEFAULT", "mode")

# Paths
# Get root directory & set as current workspace
if MODE == "PRODUCTION":
    ROOT = CONFIG.get("INPUTS", "root")
elif MODE == "TEST":
    ROOT = os.path.dirname(__file__)
else:
    err = "Unknown MODE: {}".format(MODE)
    logging.critical(err)
    raise AttributeError(err)
ROOT = ROOT.replace("/", "\\")

if not os.path.exists(ROOT):
    err = "root workspace does not exist"
    logging.critical(err)
    raise AttributeError(err)

logging.info("Execution Mode: \t{}".format(MODE))
logging.info("Workspace: \t{}".format(ROOT))

# Template folder
template_path = os.path.join(ROOT, ".template")
temp_gdb_name = "template.gdb"
logging.info("Template: \t\t{}".format(
    os.path.join(template_path, temp_gdb_name)))

# Sub-folders
sub_folders = []
if CONFIG.has_option("INPUTS", "sub_folders"):
    sub_folders = CONFIG.get("INPUTS", "sub_folders").split()
logging.info("Subfolders: \t{}".format(", ".join(sub_folders)))

# Make excel
excel_fields = CONFIG.get("INPUTS", "excel_fields").split()

if CONFIG.get("INPUTS", "excel").lower() == "true":
    MAKE_EXCEL = True
    logging.info("Excel: \t\tEnabled")
    logging.info("Excel fields: \t{}".format(", ".join(excel_fields)))
else:
    MAKE_EXCEL = False
    logging.info("Disabled: \t\tExcel")

# Set USE_ARCPY to actual boolean value
if CONFIG.get("DEFAULT", "use_arcpy").lower() == "true":
    USE_ARCPY = True
else:
    USE_ARCPY = False

# Set Spatial Reference (by name) for all output data
spatial_ref_name = CONFIG.get("INPUTS", "spatial_reference")

# Configure for TEST / USE_ARCPY modes
if USE_ARCPY is False and MODE == "PRODUCTION":
    err = "PRODUCTION Mode requires 'arcpy'"
    logging.critical(err)
    raise AttributeError(err)

if USE_ARCPY:
    logging.info("ArcPy: \t\tEnabled")
    try:
        import arcpy
        arcpy.env.overwriteOutput = False
        # Load & set spatial reference
        SR = arcpy.SpatialReference(spatial_ref_name)
        arcpy.env.outputCoordinateSystem = SR
        logging.info("Spatial Reference: {0}".format(SR.name))
    except Exception as err:
        # Log error message before aborting
        logging.critical(err)
        raise err
else:
    logging.info("ArcPy: \t\tDisabled")

# Get project names from topo maps (rasters in GDB) or test_names
logging.info("Listing projects...")
if MODE == "PRODUCTION":
    raster_gdb = os.path.join(
        ROOT,
        CONFIG.get("INPUTS", "rasters")).replace("/", "\\")
    if not os.path.exists(raster_gdb):
        err = "{} does not exist".format(raster_gdb)
        logging.critical(err)
        raise AttributeError(err)
    arcpy.env.workspace = raster_gdb
    try:
        projects = arcpy.ListRasters()
    except Exception as err:
        logging.critical(err)
        raise err
    if not projects:
        err = "No rasters found in {}".format(raster_gdb)
        logging.critical(err)
        raise AttributeError(err)
else:
    projects = CONFIG.get("DEFAULT", "test_names").split()
logging.info("Projects Found: \t{}".format(len(projects)))


# =============================================================================
# Validate & unpack dataset and feature values

logging.info("Validating datasets and featureclasses...")
logging.info(
    "\t(NOTE: 'LINE' values will automatically changed to 'POLYLINE')")
dtypes = ("POINT", "LINE", "POLYGON")
conf_err = None

# Prepare empty dict for unpacking datasets and features (list of tuples)
dataset_features = {}

for dataset in DATASETS.sections():
    # Prepare empty list for unpacked features
    dataset_features[dataset] = []
    for feature, geom_type in DATASETS.items(dataset):
        # Validate the values
        if geom_type not in dtypes:
            logging.error(
                "Invalid datatype: {} = {}".format(feature, geom_type))
            conf_err = "One or more errors with config (see log). Aborting."
        if feature == "":
            logging.warn(
                "Empty feature name: {} = {}".format(feature, geom_type))
        # Unpack
        dataset_features[dataset].append((feature, geom_type))

if conf_err:
    logging.critical(conf_err)
    raise AttributeError(conf_err)

logging.info("Datasets and featureclasses valid and prepared.")


# =============================================================================
# Functions
# =============================================================================

def make_excel(path, name):
    """Makes Excel files."""
    # Strip any extension provided
    name = os.path.splitext(name)[0]
    # Open Excel, add workbook
    try:
        excel = win32.Dispatch("Excel.Application")
        workbook = excel.Workbooks.Add()
        worksheet = workbook.Worksheets("Sheet1")
    except AttributeError as err:
        logging.critical("Please close all instances of Excel and try again.")
        raise AttributeError(err)
    # Set column headers from list in config
    for f in excel_fields:
        worksheet.Cells(1, excel_fields.index(f) + 1).Value = f
    # Save, close, and quit
    workbook.SaveAs(path + '/{0}.xlsx'.format(name))
    workbook.Close()
    excel.Quit()
    # Kill the Excel background process
    try:
        Popen("taskkill /f /im EXCEL.EXE", stdout=PIPE)
        sleep(.5)
    except:  # noqa -- Ignore naked except
        pass
    return


def create_template():
    """Creates the template project folder."""
    # Make template folder and set as workspace
    try:
        os.mkdir(template_path)
    except Exception as err:
        logging.critical(err)
        raise err
    os.chdir(template_path)
    if USE_ARCPY:
        arcpy.env.workspace = template_path

    # Make GDB, datasets, and featureclasses
    if USE_ARCPY:
        arcpy.CreateFileGDB_management(template_path, temp_gdb_name)

    else:
        os.mkdir(os.path.join(template_path, temp_gdb_name))
    logging.info("Created GDB")
    gdb_path = os.path.join(template_path, temp_gdb_name)

    # Make dataset
    for dataset in dataset_features.keys():
        logging.info("Creating {}...".format(dataset))

        dataset_path = os.path.join(gdb_path, dataset)

        if USE_ARCPY:
            arcpy.CreateFeatureDataset_management(
                gdb_path, dataset, SR)
        # or make folder
        else:
            os.mkdir(dataset_path)

        # Make features
        for feature, geom_type in dataset_features[dataset]:
            # Remove potential spaces
            feature = feature.strip()
            geom_type = geom_type.strip()

            # Validate types
            if geom_type == "LINE":
                geom_type = "POLYLINE"
            logging.info("{}/{}/{}: {}".format(
                temp_gdb_name, dataset, feature, geom_type))
            if geom_type not in ("POINT", "POLYGON", "POLYLINE"):
                err = "Illegal geometry type: {}".format(geom_type)
                logging.critical(err)
                raise IOError(err)

            if USE_ARCPY:
                try:
                    arcpy.CreateFeatureclass_management(
                        # Export to current dataset, should be workspace
                        dataset_path,
                        # Feature
                        feature,
                        # Geometry type {POINT, POLYLINE, POLYGON}
                        geom_type)
                except Exception as err:
                    logging.critical(err)
                    raise err
            else:
                f_path = "{}.txt".format(os.path.join(dataset_path, feature))
                with open(f_path, "w") as f:
                    f.write(geom_type)

    # Make sub folders
    for sub in sub_folders:
        logging.info("Creating {}/".format(sub))
        os.mkdir(os.path.join(template_path, sub))

    # Make Excel tracking file
    if MAKE_EXCEL:
        logging.info("Making Excel file...")
        try:
            make_excel(template_path, "template")
        except Exception as err:
            logging.critical(err)
            raise err
    logging.info("Template Complete.")
    return


def make_projects():
    """Copies the template for each project."""
    os.chdir(ROOT)
    for project in projects:
        logging.info("Creating {}/".format(project))
        proj_dir = os.path.join(ROOT, project)
        # Copy template and rename to project name
        shutil.copytree(template_path, proj_dir)
        # Rename template GDB to project
        os.rename(
            os.path.join(proj_dir, temp_gdb_name),
            os.path.join(proj_dir, project + ".gdb"))
        # Rename Excel file if exists
        if MAKE_EXCEL:
            os.rename(
                    os.path.join(proj_dir, "template.xlsx"),
                    os.path.join(proj_dir, project + ".xlsx"))
    return


logging.info("Script configured.")


# =============================================================================
# Main Process
# =============================================================================

# Start message
logging.info("")
logging.info("="*79)
logging.info("Starting {} {}".format(
    CONFIG.get("DEFAULT", "name"),
    CONFIG.get("DEFAULT", "version")))


# TODO: remove
'''
def main():
    for map_sheet in projects:
        # Add a blank line to the log between project folders
        logging.info(map_sheet)

        # Reset current dir to Projects folder


        # Make new folder for next map_sheet & cd into it
        sheet_name = map_sheet[0:-4]
        os.mkdir(sheet_name)
        print "Processing {0}...".format(sheet_name)
        sheet_dir = os.path.join(prjbase_dir, sheet_name)
        change_dir(sheet_dir)

        # Copy toposheet to new project folder
        logging.info("Copying {0} sheet...".format(
            os.path.join(names_dir, sheet)))
        shutil.copy2(os.path.join(names_dir, sheet), sheet_dir)

        # Make blank Excel file named for sheet_name
        logging.info("Creating xlsx file...")
        make_excel(sheet_dir, sheet_name)

        # Make GDB & cd into it
        logging.info("Creating {0}.gdb".format(sheet_name))
        arcpy.CreateFileGDB_management(sheet_dir, sheet_name)
        gdb_dir = arcpy.env.workspace + "\\" + sheet_name + ".gdb"
        arcpy.env.workspace = gdb_dir
        sleep(1)

        # Load in georeferenced raster
        logging.info(
            "Importing {0} georeferenced raster...".format(sheet_name))
        arcpy.env.workspace = georef_dir
        georef_list = arcpy.ListRasters()
        for georef in georef_list:
            if georef == sheet_name:
                georef_path = arcpy.Describe(georef).catalogPath
                arcpy.RasterToGeodatabase_conversion(georef_path, gdb_dir)
                break

        # Load in boundary
        logging.info("Importing full boundary...")
        arcpy.env.workspace = foot_dir
        bounds = arcpy.ListFeatureClasses()
        bounds.sort()
        full_boundary = arcpy.Describe(bounds[0]).catalogPath
        sheet_footprints = arcpy.Describe(bounds[1]).catalogPath
        map_elements = arcpy.CreateFeatureDataset_management(
            gdb_dir, "MapElements", SR)
        map_elements = map_elements.getOutput(0)
        arcpy.env.workspace = map_elements
        arcpy.FeatureClassToFeatureClass_conversion(
            full_boundary, map_elements, "Full_boundary")

        # Import sheet boundary
        logging.info(
            "Selecting and importing {0} footprint...".format(sheet_name))
        qry = "Name = '{}'".format(sheet_name)
        arcpy.FeatureClassToFeatureClass_conversion(
            sheet_footprints,
            map_elements,
            sheet_name + "_footprint",
            qry)

        # Create feature datasets from dataset text file names & cd into it
        for dataset_file in os.listdir(dataset_dir):
            arcpy.env.workspace = gdb_dir
            dataset_name = dataset_file[0:-4]
            logging.info(
                "Creating and populating {} dataset...".format(dataset_name))
            new_dataset = arcpy.CreateFeatureDataset_management(
                arcpy.env.workspace, dataset_name, SR)
            dataset_path = arcpy.Describe(new_dataset).catalogPath
            arcpy.env.workspace = dataset_path
            # Make feature classes from dataset textfile contents
            make_features(os.path.join(dataset_dir, dataset_file))
'''


# Run if executed directly:
if __name__ == '__main__':
    # Start the main process
    try:
        create_template()
        make_projects()
    except Exception as err:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        logging.info("Errors at line: {}".format(exc_tb.tb_lineno))
        logging.critical(err)
        raise err

    logging.info("Done.")
