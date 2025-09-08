#Program:			excel_converter.py
#Title:				New Excel Converter
#Author:			Ryan Piazza
#Date: 				09/05/2025
#-------------------------------------------------------------------
#Description		This python script wraps a folder watcher
#					around a conversion with filtering script to
#					convert both xls and xlsx (mostly) to csv.
#
#					Then it will move those to an output location
#					and then remove the original from input location
#					(along with any other files).
#-------------------------------------------------------------------
#Modification List
#Date: 09/08/2025 by Ryan Piazza
#Date: __/__/____ by _______________
#Date: __/__/____ by _______________
#Date: __/__/____ by _______________
#Date: __/__/____ by _______________
#Date: __/__/____ by _______________
#Date: __/__/____ by _______________
#-------------------------------------------------------------------

import os, time, logging, re
import pandas as pd

from watchdog.events import FileSystemEvent, FileSystemEventHandler
from watchdog.observers import Observer

#Configure logging to see events
logging.basicConfig(filename="vert_log.log",
                    level=logging.INFO,
                    format='%(asctime)s - %(message)s', 
                    datefmt='%Y-%m-%d %H:%M:%S')

#-----Functions-----

#To ensure that the file isn't being accessed before permissive
def file_wait(path, retries=10, delay=1):
    for i in range(retries):
        try:
            with open(path, "rb"):
                return True
        except PermissionError:
            time.sleep(delay)
    return False

#Keep file names POSIX compliant
def sanitize(filename: str):
    #Replace invalid characters with underscore
    return re.sub(r'[\\/*?:\[\]]', '_', filename)

#Converting to excel
def to_csv(file: str):
    try:    
        with open(file, "r"):

            #Get the filename_old for later, splitting off the file extension
            filename_old = os.path.splitext(os.path.basename(file))[0]
            logger.info(f"Old filename: {filename_old}")

            #Read the excel file
            #More information here: https://pandas.pydata.org/docs/reference/api/pandas.ExcelFile.html
            xls = pd.ExcelFile(file)

            #Get a list of the sheet names
            xls.sheet_names

            #Create a dictionary to store the sheets
            sheet_to_df_map = {}

            #Key:value pairs of sheet name to sheet df
            for sheet_name in xls.sheet_names:
                sheet_to_df_map[sheet_name] = xls.parse(sheet_name)

            #Close the Excel parser
            xls.close()

        #Take key value pairs and convert to csv for each df
        for key in sheet_to_df_map:

            filename_new = f"{filename_old}-{sanitize(key)}.csv"
            filepath_new = "output/"

            #Create the exit path by joining the new filepath ("The output folder") and the name (concat of old name and tab name)
            exit_path = os.path.join(filepath_new, filename_new)

            #Log conversion of each sheet
            logger.info(f"Converting sheet: {key} to CSV at exit path {exit_path}.")

            #Write contents of dataframe to tab separated (.csv)
            with open(exit_path, "w"):
                sheet_to_df_map[key].to_csv(exit_path, sep='\t')

    except Exception as e:
        logger.warning(f"Conversion exception: {e}")

#Fake event pitcher for items still in folder at script restart
#def existing(handler, input_path):




#-----Sysmon class object-----

#New system monitor class
class sysmon(FileSystemEventHandler):

    #Any files created...
    def on_created(self, event: FileSystemEvent):
        
        #Log the initial event and time of item moved to input folder
        logger.info(f"Created: {event.src_path}")
        
        #Pick up parsed info about the file
        base_name, extension = os.path.splitext(event.src_path)
        logger.info(f"Base name: {base_name} & Extension: {extension}")
        
        #If event-causing item exists and is a file and has one of the acceptable file extensions
        acceptable = [".xlsx", ".xls", ".csv"]
        if os.path.exists(event.src_path) and os.path.isfile(event.src_path) and (extension in acceptable):
            
            try:
                #If file is able to be accessed
                if file_wait(event.src_path):
                    to_csv(event.src_path)
                    os.remove(event.src_path)
                    logger.info(f"Removed original file: {event.src_path}")
                
                else:
                    logger.warning(f"File {event.src_path} could not be accessed after retries")

            except Exception as e:
                logger.warning(f"Conversion/removal exception: {e}")
        
        #Move files that don't belong to another folder for manual review (/manrev)
        else:
            try:
                bad_file_path = event.src_path
                filename_old = os.path.basename(bad_file_path)
                print(filename_old)
                os.rename(f"{bad_file_path}", f"{os.path.join("./manrev/", filename_old)}")
                logger.info(f"Unknown file type found in input folder: {filename_old}. Moved to manrev/ for manual review.")
            
            #Critical error: bad file squatting in input folder
            except Exception as e:
                logger.critical(f"Critical error: bad file squatting in input folder. Remove {filename_old} manually then continue program.")

            
    #Any system event involving the specified folder
    #def on_any_event(self, event: FileSystemEvent) -> None:
    #    
    #    print(event)

    #These functions could later be implemented to confirm the movement and deletion (a sort of TCP-like handshake almost)

    #Any files modified...
    #def on_modified(self, event: FileSystemEvent):
    #    
    #    print(f"Modified: {event.src_path}")
    #
    #    print(type(event.src_path))

    #Any files deleted...
    #def on_deleted(self, event: FileSystemEvent):
    #    
    #    print(f"Deleted: {event.src_path}")

    #Any files moved...
    #def on_moved(self, event: FileSystemEvent):
    #
    #    print(f"Moved: {event.src_path}")


#-----Main script body-----

if __name__ == "__main__":

    #Creating and configuring a logger
    logger = logging.getLogger()

    #Tell the observer which directory to watch
    #Use os.path.join() if not in same directory
    input_path = "./input"

    #Setting up this instance of the converter
    event_handler = sysmon()
    observer = Observer()

    #Schedule an observer thread to monitor the specified path with the custom handler
    #Recursive = true means it will also handle subdirectories
    observer.schedule(event_handler, input_path, recursive=True)

    #Start a thread that wakes up on events from the system
    observer.start()

    #Script runs at all times, constantly sleeping by 1 sec until a keyboard interrupt (Ctrl+C)
    try:
        while True:
            time.sleep(1)

    except KeyboardInterrupt:

        observer.stop()
        observer.join()
        logger.info("Observer stopped.")
