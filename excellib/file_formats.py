import win32com.client as win32
import os
import shutil
import pyeasylib
import re

def convert_xls_to_xlsx(input_file,
                        input_file_treatment="keep",
                        replace_existing_output = False,
                        backup_input_file_ext=".xlsbak", 
                        backup_folder=None):
    """
    Converts an Excel file from .xls to .xlsx format while preserving all 
    formatting.

    Args:
        input_file (str): Path to the .xls file to be converted.

        input_file_treatment (str, optional): Determines how the original 
            .xls file is handled. Accepted values are:
            - "keep": Retain the original file without any changes (default).
            - "delete": Delete the original file after successful conversion.
            - "rename_extension": Rename the original file by appending a 
              specified backup extension (see `backup_input_file_ext`) and 
              optionally move it to a backup folder (see `backup_folder`).

        replace_existing_output (bool, optional): If True, renames any 
            existing .xlsx output file by appending a timestamp before 
            overwriting. If False, generates a unique filename for the new 
            .xlsx file to avoid overwriting (default: False).

        backup_input_file_ext (str, optional): The extension to use when 
            renaming the original .xls file during backup. Ignored unless 
            `input_file_treatment` is "rename_extension". Defaults to ".xlsbak".

        backup_folder (str, optional): The directory where the renamed 
            original .xls file will be saved if `input_file_treatment` is 
            "rename_extension". If not specified, the backup file will be 
            saved in the same directory as the input file.

    Raises:
        FileNotFoundError: If the specified `input_file` does not exist.
        ValueError: If the `input_file` is not a valid .xls file.
        FileExistsError: If the backup file already exists in the target 
            directory.
        NotImplementedError: If an invalid value is provided for 
            `input_file_treatment`.

    Returns:
        str: The absolute path of the newly created .xlsx file.

    Notes:
        - This function uses the `win32com` library to interact with 
          Microsoft Excel, so Excel must be installed and accessible on the 
          system where the function runs.
        - The function automatically converts relative paths of `input_file` 
          to absolute paths for compatibility.

    Examples:
        # Basic usage: Keep the original file and create a .xlsx file
        convert_xls_to_xlsx("data.xls")

        # Delete the original file after conversion
        convert_xls_to_xlsx("data.xls", input_file_treatment="delete")

        # Rename and back up the original file in a specific folder
        convert_xls_to_xlsx(
            "data.xls", input_file_treatment="rename_extension",
            backup_folder="backups", backup_input_file_ext=".backup"
        )

        # Overwrite the existing .xlsx file if present
        convert_xls_to_xlsx("data.xls", replace_existing_output=True)
    """
    
    # Convert input_file to an absolute path - excel.Workbooks.Open only works
    #                                          with abs path
    input_file = os.path.abspath(input_file)
    
    # Get the input file directory and filename
    input_filedir, input_filename = os.path.split(input_file)
    
    # Ensure backup ext starts with .
    if not backup_input_file_ext.startswith("."):
        backup_input_file_ext = "." + backup_input_file_ext
        
    # Check input file treatment
    input_file_treatment_options = ["keep", "delete", "rename_extension"]
    input_file_treatment = input_file_treatment.lower().strip()
    if input_file_treatment not in input_file_treatment_options:
        raise Exception (
            f"Invalid input_file_treatment='{input_file_treatment}'. "
            f"Should be {input_file_treatment_options}."
            )
        
    # Ensure input file exists
    if not os.path.exists(input_file):
        raise FileNotFoundError(f"Input file '{input_file}' does not exist.")
    
    # Ensure the file has .xls extension
    if not input_file.endswith('.xls'):
        raise ValueError(f"Input file '{input_file}' is not an .xls file.")
    
    # Generate output file path by replacing .xls with .xlsx
    output_file = re.sub(".xls$", ".xlsx", input_file, flags=re.IGNORECASE) #$ end of string
    
    
    # Check if output file is already present
    if os.path.exists(output_file):
        
        # get the current timestamp
        current_ts = pyeasylib.get_current_datetime()
        
        if replace_existing_output:
            
            # rename the existing -> better not to delete
            del_output_file = output_file + f".{current_ts}.del"
            os.rename(output_file, del_output_file)
            
            #os.remove(output_file)
        
        else:
            
            # keep the original and set a new file
            output_file = pyeasylib.check_filepath(output_file)
        
    # Initialize Excel application
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False  # Keep Excel hidden during conversion
    
    try:
        # Open the .xls file
        workbook = excel.Workbooks.Open(input_file)
        
        # Save it as .xlsx
        workbook.SaveAs(output_file, FileFormat=51)  # 51 represents xlOpenXMLWorkbook (.xlsx)
        workbook.Close()
        
        # Handle original file based on backup_original parameter
        if input_file_treatment == "keep":
            
            pass
        
        elif input_file_treatment == "delete":
            
            os.remove(input_file)
        
        elif input_file_treatment == "rename_extension":
                        
            # Get the backup filename and update the extension
            backup_filename = os.path.splitext(input_filename)[0] + backup_input_file_ext
            
            # Get the correct backup folder and create if not present
            if backup_folder:
                # Create if not present
                pyeasylib.create_folder(backup_folder)
            else:
                # if not specified, save in the same folder as input
                backup_folder = input_filedir
                
            # Create the backup file
            backup_file = os.path.join(backup_folder, backup_filename)
            
            # Raise error if backup file is already present
            if os.path.exists(backup_file):
                raise FileExistsError(f"Backup file '{backup_file}' already exists.")
            
            # Rename
            shutil.move(input_file, backup_file)
            
        else:
            
            raise NotImplementedError ("Should not encounter this error.")
            
    except Exception as e:
        raise RuntimeError(f"Failed to convert file: {e}")
    finally:
        # Quit Excel application
        excel.Quit()


if __name__ == "__main__":
    
    
    if False:
        
        # Example usage
        input_file = r"./test/file4 - same as 1.xls"
    
        input_file_treatment="rename_extension"
        replace_existing_output = True
        backup_input_file_ext="xlsbak"
    
        
        backup_folder = r"./test/out1/out2"
        convert_xls_to_xlsx(input_file, 
                            input_file_treatment = input_file_treatment,
                            replace_existing_output=replace_existing_output,
                            backup_input_file_ext=backup_input_file_ext, 
                            backup_folder=backup_folder)
