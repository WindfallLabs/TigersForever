# Project Folder Setup for the Tigers Forever Project
Author: Garin Wally, March 2015 (revised & expanded March 2018)  
License: [MIT](https://choosealicense.com/licenses/mit/)


## ABOUT:
This script and related config files builds a `.template` folder with:  
* a geodatabase containing the datasets and featureclasses described in `datasets.ini`  
* an Excel tracking file  
* any other sub-folders as specified in the `config.ini` file (default: Maps, Extras)  
This template is then copied and renamed for each raster in the input raster GDB (in `config.ini`).  
Example folder hierarchy:  

* root/                 -- [1] This root folder must already exist  
    * rasters.gdb		-- [2] This geodatabase must exist containing named rasters  
    * map1/		        -- Project folders are created using raster names (above)  
        * map1.gdb	    -- GDBs, datasets and featureclasses are made using datasets.ini  
        * sub_folders/	-- [3] Subfolders to make in each project folder  
        * map1.xlsx	    -- [4] Excel tracking file  
    * map2/  
        * etc.  
        * etc.  
    * etc./  

It may look something like this:  

* C:/workspace/Parsa/  
    * D4_2784_12D_Amlekhganj/  
        * D4_2784_12D_Amlekhganj.gdb  
        * D4_2784_12D_Amlekhganj.xlsx  
        * Maps/  
        * Extras/  
    * etc./  
        * etc.  


## SETUP:
To prepare for use, open the config.ini file in a text editor.  
This file is the main configuration file for the script.  
Do not change the section names ([SECTION]) or the "options" (names before the "=").  
You can, and in some instances must, change the values after the "=".  

Ensure that the mode is set to PRODUCTION (TEST is commented out)  
e.g.  

    #mode = TEST
    mode = PRODUCTION

[1] Set the root path (e.g. <path-to-whatever>).  

[2] Set the name of the existing GDB containing georectified rasters that will
be used to create each project folder in root.  

    rasters = rasters.gdb

**TODO:** I think the full boundary and or fishnet/grid system should also be in here  

[3] Set which subfolders should be made in each project folder (seperated by spaces)  

    sub_folders = Maps Extras

[4] Turn on/off Excel file creation using the true or false options.  

    #excel = False
	excel = True

And set which fields to create (seperated by spaces)  

    excel_fields = NAME DATE TIME TASKS

