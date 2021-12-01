# Copyright (c) 2020 - owgs
# License: Open-source MIT

'''
Wrapper functions for the openpyxl library.
'''

# Import dependencies from py standard libs
import re

# Import dependencies from openpyxl
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font


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
        
    
if __name__ == "__main__":
    
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