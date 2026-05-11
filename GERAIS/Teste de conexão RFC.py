from pyrfc import get_nwrfclib_version, __version__
import os
print("PyRFC:", __version__)
print("SAPNWRFC_HOME:", os.environ.get("SAPNWRFC_HOME"))
print("NWRFC SDK:", get_nwrfclib_version())