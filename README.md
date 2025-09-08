# Improved Excel Converter

## Rationale 
The current Excel converter in use is very good, but several years out of date. Any conversions requiring something more complex than changing the file extension aren't really covered by the converter. 

The conversion to .csv additionally will not support multi-page client files, as Excel only allows one worksheet per CSV. This means only the last sheet in a multi-page .xlsx or .xls file would survive, potentially excising important customer data. 

## The Structure Implemented 

The python script combines two powerful libraries to create an efficient file conversion process with the potential to be extended in many ways: watchdog (a system event monitor) and pandas (a data frame handling library). 

The system monitor watches for the creation of files in the "input" directory, then processes them dependent on their type. 

## The Build

The build was done with Pyinstaller and removes the need to implement a Python interpreter on the endpoint the script is used on. Updates to source code will not be built into the executable until compiled again after pull, so ensure that this is done before updating in production. 

The build is located in the dist/ directory. 

## Dependencies

