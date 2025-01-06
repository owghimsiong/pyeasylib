# Version
__version__ = "Version 1.0 (last updated 2023-02-09)"

from pyeasylib.base import *
from pyeasylib.base_datetime import *
import pyeasylib.pdlib as pdlib
import pyeasylib.excellib as excellib

# Added a try except block for this, in case this is not used and the 
# dependencies are not installed.
try:
    import pyeasylib.dblib as dblib
except ModuleNotFoundError as e:
    print (f"ModuleNotFoundError: pyeasylib.dblib not imported as: {str(e)}")
    pass