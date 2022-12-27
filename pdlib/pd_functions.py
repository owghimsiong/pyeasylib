# -*- coding: utf-8 -*-
"""
Contains wrapper functions for pandas classes.
"""

import pandas as pd
import numpy as np
import logging
import datetime as dt
import time
import os
from collections import OrderedDict
from pyeasylib.base import pretty_list

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

def assert_unique_series_mapping(
    series, 
    assert_error_message = None,
    ):
    '''
    Makes sure that one index only returns one value.
    
    Otherwise, will raise an error.
    '''
    
    # Check if there are any duplicated indices. If there are,
    # means that that index maps to more than one values.
    # keep = False -> return all duplicates
    duplicated = series.index.duplicated(keep = False)
    
    # Raise error if there are duplicates
    if duplicated.any():
        
        # Create a dataframe of duplicates 
        dup_df = series[duplicated].to_frame()
    
        # Get the error message
        if assert_error_message is None:
            
            val_name = "values" if series.name is None else series.name
            idx_name = "index" if series.index.name is None else series.index.name
            assert_error_message = f"Duplicated {val_name} found for {idx_name}."
        # 
        error = (
            f"{assert_error_message}\n"
            f"{dup_df.__repr__()}"
            )
        logger.error(error)
        raise Exception (error)
        
    else:
        
        return series
    
def get_key_to_values(
    df, key_column, value_column,
    assert_one_to_one = True,
    assert_error_message = None
    ):
    '''
    
    assert_error_message -> set automatically if None. 
    
    See __assert_unique_series_mapping__.
    '''
    
    # Get the key and values matrix
    key_to_value_series = (
        df[[key_column, value_column]]
        .drop_duplicates()
        .set_index(key_column)[value_column]
        .sort_index()
        )
    
    # Check if there are invoices with multiple amts
    # Will raise error if there's duplicated indices.
    if assert_one_to_one:
        
        assert_unique_series_mapping(
            key_to_value_series,
            assert_error_message)
    
    return key_to_value_series

def get_main_table_from_df(df0, expected_header_columns,
                           return_only_expected_header_columns = False):
    '''
    Extract a main data table from the raw data loaded as dataframe.
    
    For example, this is used when the data contains multiple header
    rows that are not part of the main table.

    Parameters
    ----------
    df0 : DATAFRAME
        The data to be filtered.
    
    expected_header_columns : LIST
        This specifies the header columns of the header. This row is
        considered the header and subsequent rows will be set as the
        data row.
    
    return_only_expected_header_columns : BOOLEAN, (default False)
        If set to True, the output will only contains the columns
        specified in expected_header_columns.
        
    Raises
    ------
    Exception
        DESCRIPTION.

    Returns
    -------
    df : DATAFRAME
        Dataframe of the extracted data
    '''

    # Take the unique set of the headers.
    expected_header_set = set(expected_header_columns)
    
    # Find the header row
    header_idx = None
    for i in range(df0.shape[0]):
        
        values = df0.iloc[i]
        intersection = expected_header_set.intersection(values)
        if intersection == expected_header_set:
            
            header_idx = i
            break
    
    # Raise error if cannot find header row
    if header_idx is None:
        error = (
            "Unable to identify header row in file. "
            f"Header must contain the following columns: {expected_header_columns}."
            )
        logger.error(error)
        raise Exception (error)
    
    # Set the dataframe
    df = df0.iloc[(header_idx+1):]
    df.columns = df0.iloc[header_idx].values
    
    # Drop NA because sometimes there will be s/n entries although no 
    # actual data
    df = df.dropna(subset = expected_header_columns, how='all')

    # keep only required columns
    if return_only_expected_header_columns:
        df = df[expected_header_columns]
        
    return df

def series_to_duplicates(s, **kwargs):
    '''
    quick function to extract the duplicated data from
    a series input.

    input:
        s - series, list, tuple or array etc.
        kwargs - arguments to prettyList
                 - max_ends: default 10,
                 - linker  : default '...'
                 - sep     : ','
                 - last    : 'and'

    returns:
        tuple of (num_duplicates, series of duplicates, dup_str)
    '''

    s = pd.Series(list(s))                            
    counts = s.value_counts(dropna=False)
    dup_counts = counts[counts>1]

    num_dups = len(dup_counts)

    # get parameters
    dup_str = series_to_count_string(dup_counts, **kwargs)
    final = "Number with count>1: %s - %s." % (num_dups, dup_str)

    return num_dups, dup_counts, final

def series_to_count_string(s, **kwargs):
    '''
    Takes in a series, and return a string of counts where
    such as index1(x4), index2(x5).

    Note: This function does not count the number of values, i.e.
          does not do .value_counts().
          This function will just list the index and the value associated
          with that index.
          Expected to be used if the series is already a series of counts.

    **kwargs are passed to prettyList.
    '''

    #added try/except block to take into account unexpected
    #data types
    try:
        data_list = ["%s(x%s)" % i for i in s.items()]
    except AttributeError as e:
        error = "%s. Input s is type=%s but must be pd.Series." % \
            (e, type(s))
        raise AttributeError(error)

    # get parameters
    max_ends = kwargs.get('max_ends', 10)
    linker = kwargs.get('linker', '...')
    sep = kwargs.get('sep', ', ')
    last = kwargs.get('last', 'and')
    count_string = pretty_list(data_list,
                               max_ends=max_ends,
                               linker=linker,
                               sep=sep,
                               last=last)
    return count_string


def validate_columns_exist(df, columns, optional = False):
    '''
    Check if columns are present in the input dataframe.
    
    If optional is set to True, columns that are set to None will 
    be excluded from the checks.
    '''
    
    # filtered columns
    if optional:
        columns = [c for c in columns if c is not None]
    
    # check if exists
    not_found = [c for c in columns if c not in df.columns]
    
    # raise error
    if len(not_found) > 0:
        
        # format error msg
        err = (
            f"Columns(s) not found: {not_found}."
            "\n\n"
            f"{df.head().__repr__()}"
            )
        
        # raise
        raise LookupError (err)

if __name__ == "__main__":
    
    
    pass
