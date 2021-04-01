"""
Contains useful functions.
"""

# Import packages from standard libraries
import pandas as pd
import numpy as np
import logging
import datetime as dt
import time
import os
from collections import OrderedDict

# Initialise logger
logger = logging.getLogger()
if not(logger.hasHandlers()):
    
    # Define and format stream handlers
    stream_handler = logging.StreamHandler()
    stream_formatter = logging.Formatter('[%(levelname)s] %(message)s')
    stream_handler.setFormatter(stream_formatter)
    stream_handler.setLevel(logging.DEBUG)
    
    # Add the logger
    logger.addHandler(stream_handler)
    logger.setLevel(logging.DEBUG)


################## Useful functions ##################
    
def create_folder(folderpath,
                  alternate_if_exists = False):
    '''
    Creates folder. Will recursively create folder till the folderpath
    is created.     
    
    Parameters
    ----------
    folderpath: str
        Folderpath to create in string format
        
    alternate_if_exists: boolean (default False)
        Whether to create the folder if it is already present.
        
        False : Do not create alternate folder
        True  : Create alternate folder if it is already present and
                not empty.

    Returns
    -------
        folderpath
    
    CHANGES:
    #20200818 - initialised by owgs.
    '''
    
    if not(os.path.exists(folderpath)):
        
        try:
            os.mkdir(folderpath)
        
            logger.info("Created folder=%s." % folderpath)
            
            return folderpath
        
        except FileNotFoundError:
        
            parent_folder_path, new_foldername = os.path.split(folderpath)
            
            # Create parent folder path
            create_folder(parent_folder_path, alternate_if_exists = alternate_if_exists)
            
            # then create the folder path
            create_folder(folderpath, alternate_if_exists = alternate_if_exists)
            
            return folderpath
    
    else:
        
        if (alternate_if_exists is True) and (len(os.listdir(folderpath)) > 0):
            
            folderpath = "%s_%s" % (folderpath, get_current_datetime())
            
            # then create the folder path
            create_folder(folderpath, alternate_if_exists = alternate_if_exists)

            return folderpath
        
        else:
            
            return folderpath

def get_current_datetime():
    '''
    Returns the current date time in YYYYMMDDHHMMSS format.

    Returns
    -------
    str
        current date time in YYYYMMDDHHMMSS
    '''
    
    return time.strftime("%Y%m%d%H%M%S")

def check_filepath(fp):
    '''
    This function returns a filepath that does not exist yet.
    
    This is typically used when you want to generate an output
    file, but do not wish to overwrite any file.
    
    Therefore, this function ensures that if the desired fp 
    is already present, then it will add a datetimestamp
    to force a new fp.
    
    Otherwise, it will return the fp that was provided to this 
    function.
    '''
    
    if not os.path.exists(fp):
               
        return_fp = fp

    else:
        
        # Get the components
        dirname, file_name_ext = os.path.split(fp)
        
        # Get the folderpath
        if dirname == "":
            dirname = os.getcwd()
                    
        # get the filename and extension
        file_name, file_ext = os.path.splitext(file_name_ext)
        
        while True:
            
            # Get the current datetimestamp
            datetimestamp = get_current_datetime()
            
            # create the new file
            new_file_name_ext = file_name + "_" + datetimestamp + file_ext
            
            new_fp = os.path.join(dirname, new_file_name_ext)
            
            if os.path.exists(new_fp):
                
                pass
            
            else:
                
                return_fp = new_fp
                break
            
    return return_fp    

def assert_one_and_get(inlist, what="items"):
    '''
    This function checks whether there's only one element in 
    the list.   
    '''
    
    if len(inlist) > 1:
        raise Exception ("Multiple %s found: %s." % (what, inlist))
    elif len(inlist) == 0:
        raise Exception ("No %s found." % what)
    else:
        return inlist[0]
           

def assert_unique(data, what = "data"):
    '''
    Checks the uniqueness of the input data.
    '''
    
    series = pd.Series(data)
    counts = series.value_counts().sort_index()
    
    # dup counts
    dup_counts = counts[counts > 1]
    
    # Raise error if there's duplicates
    if dup_counts.shape[0] > 0:
        
        dup_str = "; ".join([f"{k} (x{v})" for k, v in dup_counts.items()])
        msg = f"Duplicated {what} found: {dup_str}."
        
        logger.error(msg)
        raise Exception (msg)


def ordered_unique(in_list):
    
    return list(OrderedDict.fromkeys(in_list))


def get_filepaths_for_folder(input_folder, 
                             valid_extensions = [],
                             deep = False):
    '''
    Returns a list of excel filepaths from an input folder.
    
    This function can scan in deeper folders to extract all files
    that satisfy the file extension criteria.
    
    Input:
        input_folder: str
            Folder path in full
            
        valid_extensions: list
            A list of all valid file extensions
            
        deep: boolean
            Specifies if the search for the filepaths will go beyond
            the input_folder.
            
            False: Default; only filepaths from the input_folder
                   will be extracted
                   
            True:  Filepaths from folder nested within the input
                   folder will also be returned

    Returns:
        a list of filepaths                   
    '''
    
    # normalise the valid extensions to lower case and add a dot in front
    valid_extensions = [i.lower() for i in valid_extensions]
    valid_extensions = [i if i.startswith(".") else ".%s" % i
                        for i in valid_extensions]

    # Get the dir entries 
    entries =  os.listdir(input_folder)
            
    # Get the files
    filenames = [f for f in entries 
                 if os.path.isfile(os.path.join(input_folder, f))]
    dirnames = [f for f in entries if f not in filenames]
        
    # Get only excel files and not folders and skip temp files
    filenames = [f 
                 for f in filenames
                 if (not(f.startswith("~")) and 
                     os.path.splitext(f)[-1].lower() in valid_extensions)] 
    
    # calculate the filepaths
    filepaths = [os.path.join(input_folder, fn) for fn in filenames]
    
    # Get for the folder
    if deep and len(dirnames) > 0:
        
        for dirname in dirnames:
            dirpath = os.path.join(input_folder, dirname)
            filepaths2 = get_filepaths_for_folder(
                dirpath,
                valid_extensions=valid_extensions,
                deep = deep)
            
            # update
            filepaths.extend(filepaths2)
               
    return sorted(list(set(filepaths)))
     
if __name__ == "__main__":
    
    
    pass