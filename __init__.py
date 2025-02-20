# Version
__version__ = "2.1"
__version_date__ = "20 Feb 2025"

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
