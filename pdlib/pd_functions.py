# -*- coding: utf-8 -*-
"""
Created on Sat Mar 13 13:45:56 2021

@author: OwGhimSiong
"""

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


def get_table_from_df(df0, expected_header_columns,
                      return_only_expected_header = False):
    '''
    Extract a table properly.
    '''
    
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
    if return_only_expected_header:
        df = df[expected_header_columns]
        
    return df


     
if __name__ == "__main__":
    
    
    pass