import cx_Freeze
import os
import sys

os.environ["TCL_LIBRARY"] = r"C:\LOCAL_TO_PYTHON\Python35-32\tcl\tcl8.6"
os.environ["TK_LIBRARY"] = r"C:\LOCAL_TO_PYTHON\Python35-32\tcl\tk8.6"

base = None

if sys.platform =="win32":
    base = "Win32GUI"
includefiles = ['*.js', '*.json']
includes = []

cx_Freeze.setup(
    name= "a",
    version = "0.1",
    options = {"build_exe": {"includes": ["selenium"]}},
    executables = [cx_Freeze.Executable("IDMS_delegation-option1.py", base = None)]
)