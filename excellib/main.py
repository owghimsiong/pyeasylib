# Copyright (c) 2020 - owgs
# License: Open-source MIT

'''
Wrapper functions for the openpyxl library.
'''

# Import dependencies from py standard libs
import re

import pandas as pd

# Import dependencies from openpyxl
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
utils = openpyxl.utils

# Start of class

class CellRangeIterator:
    '''
    Cell iterator to format cells easily.
    #CHANGES
    #20200422 - initialised by owgs    
    '''
    
    def __init__(self, cell_range, ws):
        
        # Get the anchors
        anchors = cell_range.upper().split(":")
        num_anchors = len(anchors)
        if num_anchors == 1:
            anchor1 = anchors[0]
            anchor2 = anchor1
        elif num_anchors == 2:
            anchor1, anchor2 = anchors
        else:
            raise Exception ("Invalid cell range: %s" % cell_range)
        
        # Split the anchors
        row1, col_alpha_1 = self.__split_anchor__(anchor1)
        row2, col_alpha_2 = self.__split_anchor__(anchor2)
        
        # reorder the rows
        rows= sorted([row1, row2])
        
        # reorder the columns
        col_alphas = sorted([col_alpha_1, col_alpha_2])
        cols = [
            openpyxl.utils.cell.column_index_from_string(col_alphas[0]),
            openpyxl.utils.cell.column_index_from_string(col_alphas[1])
            ]

        # Save to self
        self.ws = ws
        self.cell_range = cell_range
        self.rows = rows
        self.cols = cols
        self.col_alphas = col_alphas
        
        
    def iterate_over_range(self):
        '''
        for cell in self.iterate_over_range():
            
            # do something with cell
            ... ...            
        '''       
        
        # Access the rows and cols
        min_row, max_row = self.rows
        min_col, max_col = self.cols
        
        # Generator
        for ridx in range(min_row, max_row+1):
            
            for cidx in range(min_col, max_col+1):
                
                yield self.ws.cell(ridx, cidx)
                
                
    def apply_cell_function_over_range(self, cell_function):
        '''
        Applies a custom function to modify the range.
        Compare with apply_cell_format_over_range. 
        
        For example:
        
            def cell_function(cell):
                cell.value = 1; 
                cell.font = Font(color="008000")
        
            self = CellRangeIterator("A1:C10", ws)
            self.apply_cell_function_over_range(cell_function)
        '''
        
        for cell in self.iterate_over_range():
            
            cell_function(cell)
    
    
    def apply_cell_format_over_range(self, cell_attribute_dict):
        '''
        Applies formatting to cell based on setting attributes.
        Compare with apply_cell_function_over_range.
        
        For example:
            cell_range3 = "H2:J10"
            self = CellRangeIterator(cell_range3, ws)
            self.apply_cell_format_over_range({"value": 2, 
                                                "font": Font(color='008000')
                                                })
        '''
        
        for cell in self.iterate_over_range():
            
            for cell_attr, cell_attr_val in cell_attribute_dict.items():
                
                setattr(cell, cell_attr, cell_attr_val)
        
        
    def __split_anchor__(self, anchor):
        '''
        split anchor = A10 to A and 10
        
        Returns row and column.
        '''
        
        # Split to extract the alpha and digit
        reg_exp = '\D+|\d+'
        extracted = re.findall(reg_exp, anchor)
        num_extracted = len(extracted)
        
        # Get the row and column for this anchor
        column = None; row = None
        if num_extracted == 2:
            
            column, row = extracted
                    
        elif num_extracted == 1:
            
            if extracted[0].isalpha():
                column = extracted[0]
            else:
                row = extracted[0]
        
        else:
            
            raise Exception ("Unexpected anchor: %s." % anchor)
        
        # Format row to int
        try:
            
            row = int(row)
        
        except ValueError:            
        
            raise Exception ("Unexpected anchor: %s." % anchor)
                        
        return row, column
        
        
def format_cell_to_2_decimals_or_dash(cell):
    '''
    Format to 2 decimals, or to dash if 0.
    '''
    
    cell.number_format = '0.00;-0.00;"-";@'
        
def convert_df_to_excel_rows_cols(df0):
    '''
    assume df0 index and columns starts from 0.
    '''
    
    def __check_0_start_and_contiguous__(index):
        
        # check 0
        if index[0] != 0:
            raise Exception (
                f"df index/columns does not start from 0: {index}"
                )
        
        # Check contiguous
        expected_vals = range(index.shape[0])
        if list(expected_vals) != index.values.tolist():
            raise Exception (
                f"df index/columns does not have contiguous range: {index}"
                )

    # make a copy to avoid changing the original input
    df = df0.copy()
    
    if df.shape[0] > 0:
            
        __check_0_start_and_contiguous__(df.index)
        __check_0_start_and_contiguous__(df.columns)
        
        # Update to xl rows and cols    
        df.index = df.index + 1
        df.index.name = "ExcelRow"
    
        df.columns = (df.columns + 1).map(utils.get_column_letter)
        df.columns.name = "ExcelCol"
    
    return df    
    

def read_excel_with_xl_rows_cols(fp, sheet_name = 0):
    '''
    Read Excel sheets but force rows and cols to Excel alphanum.
    '''
    
    data = pd.read_excel(
        fp, 
        sheet_name = sheet_name,
        index_col = None,   # Hardcoded 
        header = None,      # Hardcoded
        skiprows = 0        # Hardcoded
        )
    
    # CHeck type
    type_data = type(data)
    if type_data is pd.DataFrame:
        output = convert_df_to_excel_rows_cols(data) 
    elif type_data is dict:
        output = {}
        for sheet, df in data.items():
            output[sheet] = convert_df_to_excel_rows_cols(df)
    else:
        raise TypeError (
            f"Unexpected format: {type_data}"
            )
    return output
    
def range_excel_letters(start, stop, include_right=False, step=1):
    """
    Get a range of Excel column alphabets from the start and stop 
    alphabet.
    
    If include_right = True, the stop alpha will be included.
    
    Usage:
        > range_excel_letters("A", "C")
        Out[4]: ['A', 'B']
        
        > range_excel_letters("A", "C", True)
        Out[5]: ['A', 'B', 'C']
        
        > range_excel_letters("A", "C", True, 2)
        Out[6]: ['A', 'C']
        
        > range_excel_letters("A", "C", False, 2)
        Out[7]: ['A']
    """ 
    
    # conditions
    if not start.isalpha():
        raise Exception (f"start is not a letter: {start}.")
        
    if not stop.isalpha():
        raise Exception (f"stop is not a letter: {stop}.")
        
    # add delta, if include right is True
    delta = 1 if include_right else 0
    
    # start_index
    start_idx = utils.column_index_from_string(start)
    stop_idx = utils.column_index_from_string(stop)
    
    # Get the data    
    output_list = []
    for idx in range(start_idx, stop_idx+delta, step):
        output_list.append(utils.get_column_letter(idx))
        
    return output_list

if __name__ == "__main__":

    # Read excel file with excel rows and cols
    fp = r"D:\Desktop\owgs\CODES\rsmsglib\fifo\aging.xlsx"
    df = read_excel_with_xl_rows_cols(fp, sheet_name = "Movement")

    
    
    # Testing the class for CellRangeIterator
    if False:
        
        # initialise the workbook
        wb = openpyxl.Workbook()
        
        # Get the active sheet
        ws = wb.active
        ws.title = "First sheet"
        
        # For cell range A1:B10
        cell_range = "A1:B10"
        self = CellRangeIterator(cell_range, ws)
        for c in self.iterate_over_range():
            c.value = str(c)
            c.font = Font(color='008000')
        
        # For cell range F2:D90
        cell_range2 = "F2:D90"
        def cell_function(cell):
            cell.value = 1
            cell.font = Font(color="008000")
        self2 = CellRangeIterator(cell_range2, ws)
        self2.apply_cell_function_over_range(cell_function)
        
        # For cell range H2:J10
        cell_range3 = "H2:J10"
        self = CellRangeIterator(cell_range3, ws)
        self.apply_cell_format_over_range({"value": 2, 
                                            "font": Font(color='008000')
                                            })
        #wb.save('test_output.xlsx')