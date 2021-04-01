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

# get float types and int types - may be different in different np versions
FLOATTYPES = set([
    np.__getattribute__(i)
    for i in ['float', 'float16', 'float32', 'float64', 'float96', 'float_']
    if i in dir(np)
    ])
INTTYPES = set([
    np.__getattribute__(i)
    for i in ['int', 'int0', 'int8', 'int16', 'int32', 'int64', 'int_']
    if i in dir(np)
    ])
NUMERICTYPES = FLOATTYPES.union(INTTYPES)

################## Useful functions ##################
def get_filepaths_for_folder(input_folder, valid_extensions = [],
                             deep = False):
    '''
    Returns a list of excel filepaths.
    
    valid_extensions = a list, or EXTENSIONS_EXCEL
    
    if deep = True, will go down into internal folders to extract.
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
        
def create_folder_for_filepath(filepath):
    '''
    Create folders for a filepath.
    
    Sometimes when we are saving a file to C:\A\B\C\out.txt, there
    will be an error if the folder C:\A\B\C is not present.
    
    This function helps to build the parent folders.
    '''
    
    # Get the dirpath
    dirpath = os.path.dirname(filepath)
    
    # Create
    outdirpath = create_folder(dirpath, alternate_if_exists = False)
    
    return outdirpath

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

def assert_no_duplicates(data, index=False, column=None):
    '''
    Check duplicates for a one-dimensional data such as an input series,
    list, dataframe index or dataframe columns.
    
    Input:
        data - list, series
               or dataframe (index or column need to be set)
        index - boolean
                only used when data is a dataframe.
                Set True, to assess uniqueness of index
        column - column name
                 only used when data is a dataframe.
                 Set column name of the dataframe to assess the uniqueness
                 of that column.
    
    If there are duplicates, exception will be raised.    
    '''
    
    def __count_and_dup_values__(values):
        
        # count occurrence and get duplicates
        value_counts = values.value_counts() 
        value_with_dup_counts = value_counts[value_counts > 1].index.tolist()
        
        return value_with_dup_counts
                
    type_data = type(data)
    if type_data is pd.DataFrame:
        
        if (index is False) and (column is None):
            raise Exception ("If index is False, you must provide a column name.")
        
        if (index is True) and (column is not None):
            raise Exception ("If index is True, column name must be None.")
            
        # get value
        values = data.index if index else data[column]
        value_with_dup_counts = __count_and_dup_values__(values)
        
        # get dup data
        dup_data = data.loc[value_with_dup_counts] if index  \
                   else data[data[column].isin(value_with_dup_counts)]
        if index:
            dup_data = dup_data.sort_index()
        else:
            dup_data = dup_data.sort_values(column)
    
        if dup_data.shape[0] > 0:
            key = 'dataframe index' if index else ("column = '%s'" % column)
            raise Exception ("Duplicates found for %s:\n\n%s" % \
                             (key, dup_data.__repr__()))
    
    elif type_data is pd.Series:
        
        # get values
        values = data.index if index else data.values
        value_with_dup_counts = __count_and_dup_values__(values)

        # count 
        dup_data = data.loc[value_with_dup_counts] if index  \
                   else data[data.values.isin(value_with_dup_counts)]
                   
        if dup_data.shape[0] > 0:
            key = 'series index' if index else 'series values'
            raise Exception ("Duplicated found for %s:\n\n%s" % \
                             (key, dup_data.__repr__()))
                
    elif type_data in [list, tuple, np.array, pd.Index]:
        
        # get and count
        values = pd.Series(data)
        value_with_dup_counts = __count_and_dup_values__(values)
        
        if len(value_with_dup_counts) > 0:
            raise Exception ("Duplicated values found: %s." % value_with_dup_counts)
        
    else:
        
        msg = "Input data is not df, series, list, tuple or arrays. Was %s." % type_data
        
        raise NotImplementedError (msg)
        
def summarise_list(
    inputlist,
    max_ends=10,
    linker='...'):
    '''
    summarize a long list into a shortened list for presentation purposes only
    '''
    
    # check type first
    inputtype = type(inputlist)
    if inputtype in [list, tuple, set, np.ndarray]: #added set on 20140526
        inputlist = list(inputlist)                 #20160331 - added np.ndarray
    else:
        error = "Invalid input type: %s." % inputtype
        logger.error(error)
        raise Exception(error)

    # return empty string if list is empty
    if inputlist == []:
        return []

    # max_ends must be positive integer
    if max_ends <= 0:
        msg = "Invalid max_ends = %s. Must be positive integer." % max_ends
        logger.error(msg)
        raise Exception(msg)

    # get outputlist
    inputlength = len(inputlist)
    max_num = max_ends * 2
    if inputlength <= max_num:
        outputlist = inputlist
    else:
        outputlist = inputlist[:max_ends] + \
                    [linker] + \
                    inputlist[(-1*max_ends):]

    return outputlist

def pretty_list(
    inputlist,
    max_ends = 10,
    linker = '...',
    sep = ', ',
    last = 'and',
    ):
    '''
    Convert a long list of items to a summarized list.
        - first part shortens the list based on the number of ends, and
          reduce the middle portion to a linker string.
        - second part convert the condensed list to a string
          with elements separated by sep, and ending with last.

    Usage:
        >>> prettyList(range(100), max_ends=5)
        '0, 1, 2, 3, 4, ..., 95, 96, 97, 98 and 99'
        >>> prettyList(['one'])
        'one'
    '''

    #block removed and shifted to a new fn summarize list
    #call that fn here
    outputlist = summarise_list(inputlist, max_ends=max_ends, linker=linker)
    if outputlist == []:
        return ''

    # convert elements to string
    outputlist = list(map(str, outputlist))

    # convert to full string
    if len(outputlist) == 1:
        outputstr = outputlist[0]
    else:
        outputstr = '%s %s %s' % \
            (sep.join(outputlist[:-1]), last, outputlist[-1])

    # return
    return outputstr


if __name__ == "__main__":
    
    
    pass
