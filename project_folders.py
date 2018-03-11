# -*- coding: utf-8 -*-
"""
Project folder setup script for Kevin McManigal and the Tigers Forever
mapping project
Author: Garin Wally
Date: March 7, 2015 (heavily revised March 2018)
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
import win32com.client as win32
from ConfigParser import SafeConfigParser
from subprocess import Popen, PIPE
from time import ctime, sleep

#import arcpy


# =============================================================================
# Setup Log File
# =============================================================================

# Time format
#t_fmt = "%H:%M:%S"

# Output log file & message format
log_name = dt.datetime.now().strftime("logs/log-%Y-%m-%d_%H.%M.txt")
logging.basicConfig(filename=log_name,
                    filemode = 'w',
                    #format='%(levelname)s: %(message)s',
                    #format='%(message)s',
                    format='[%(asctime)s] %(levelname)-8s : %(message)s',
                    level=logging.DEBUG)

logging.info("Logger ready.")


# =============================================================================
# Configuration
# =============================================================================

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

# Validate dataset and feature values
logging.info("Validating datasets.ini...")
dtypes = ("POINT", "LINE", "POLYGON")
conf_err = None
for dataset in DATASETS.sections():
    for feature, geom_type in DATASETS.items(dataset):
        if geom_type not in dtypes:
            logging.error(
                "Invalid datatype: {} = {}".format(feature, geom_type))
            conf_err = "One or more errors with config (see log). Aborting."
        if feature == "":
            logging.warn(
                "Empty feature name: {} = {}".format(feature, geom_type))
            
if conf_err:
    logging.critical(conf_err)
    raise AttributeError(conf_err)

logging.info("Configured")

logging.info("Starting {} {}".format(
    CONFIG.get("DEFAULT", "name"),
    CONFIG.get("DEFAULT", "version")))


# TODO: continue line-by-line review from here
if CONFIG.get("INPUTS", "excel").lower() == "true":
    make_excel
    
MODE = CONFIG.get("DEFAULT", "mode")

if MODE == "PRODUCTION":
    arcpy.env.overwriteOutput = False
    # List topo sheet jpg names
    arcpy.env.workspace = names_dir
    logging.info("Listing topos...")
    topos = arcpy.ListRasters()

root = CONFIG.get("INPUTS", "root")

# Set directory containing datasets and their features
dataset_dir = os.path.join(root, "Datasets")

# Set directory of original topo sheet (.tif) names
names_dir = os.path.join(root, "ClippedParsaTopos")


# Set Spatial Reference (by name) for all output data
spatial_ref_name = CONFIG.get("INPUTS", "spatial_reference")

try:
    if MODE == "PRODUCTION":
        SR = arcpy.SpatialReference(spatial_ref_name)
        arcpy.env.outputCoordinateSystem = SR
    logging.info("Spatial Reference set: {0}".format(SR.name))
except Exception as e:
    logging.error(e)
    raise e

logging.info("Paths & vars set")


# =============================================================================
# Functions
# =============================================================================

def change_dir(path):
    """Changes current directory in both arcpy and os (for safety)."""
    arcpy.env.workspace = path
    os.chdir(path)


def make_excel(path, name):
    """Makes Excel files."""
    # Get field list from config
    excel_fields = CONFIG.get("STRUCTURE", "excel_fields").split()
    # Open Excel, add workbook
    excel = win32.Dispatch("Excel.Application")
    workbook = excel.Workbooks.Add()
    worksheet = workbook.Worksheets("Sheet1")
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
    except:
        pass


def make_features(dataset_file):
    """Creates the datasets and features in datasets.ini config file."""
    for dataset in DATASETS.sections():
        if MODE == "PRODUCTION":
            arcpy.CreateFeatureDataset_management(
                arcpy.env.workspace, dataset, SR)
        
        for feature_name, geom_type in DATASETS.items(dataset):
            if MODE == "PRODUCTION":
                arcpy.CreateFeatureclass_management(
                    # Export to current workspace
                    arcpy.env.workspace,
                    # Feature
                    feature_name,
                    # Geometry type {POINT, LINE, POLYGON}
                    geom_type)
            else:
                pass
    return


# =============================================================================
# Main Process
# =============================================================================
# Create project directory schemes for each topo

def main():
    for sheet in topos:
        # Add a blank line to the log between project folders
        logging.info("")
        
        # Reset current dir to Projects folder
        change_dir(prjbase_dir)
        
        # Make new folder for next sheet & cd into it
        sheet_name = sheet[0:-4]
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
        logging.info("Importing {0} georeferenced raster...".format(sheet_name))
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
        logging.info("Selecting and importing {0} footprint...".format(sheet_name))
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
        

# Run if executed directly:
if __name__ == '__main__':
    # Start the main process
    try:
        main()
    except Exception as e:
        logging.info("Errors:")
        logging.error(e)
    
    # End of script messages
    end_time = ctime()
    end_msg = "\nScript ended: {0}".format(end_time)
    logging.info(end_msg+'\n')
