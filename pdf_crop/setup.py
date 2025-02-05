import sys
from cx_Freeze import setup, Executable
import os
import msilib


# Set up paths for TCL and TK libraries
PYTHON_INSTALL_DIR = os.path.dirname(os.path.dirname(os.__file__))
os.environ['TCL_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tcl8.6')
os.environ['TK_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tk8.6')

# Base for Windows GUI
base = None
if sys.platform == 'win32':
    base = 'Win32GUI'


# Increase recursion limit
sys.setrecursionlimit(3000)


# Executable definition
executables = [
    Executable(
        'stp.py',
        icon="logo.ico",
        base=base,
    )
]

# MSI Shortcut table
shortcut_table = [
    (
        "DesktopShortcut",  # Shortcut
        "DesktopFolder",  # Directory_
        "Pdf Crop",  # Name
        "TARGETDIR",  # Component_
        "[TARGETDIR]stp.exe",  # Target
        None,  # Arguments
        None,  # Description
        None,  # Hotkey
        None,  # Icon
        None,  # IconIndex
        None,  # ShowCmd
        "TARGETDIR",  # WkDir

        
    )
]

# MSI Options
msi_data = {"Shortcut": shortcut_table}
bdist_msi_options = {
    "data": msi_data}

# Build Options
options = {
    "bdist_msi": bdist_msi_options,
    "build_exe": {
        "include_files": [
            os.path.join(PYTHON_INSTALL_DIR, "DLLs", "tk86t.dll"),
            os.path.join(PYTHON_INSTALL_DIR, "DLLs", "tcl86t.dll"),
            "logo.ico",  # Application icon

        ],
        "packages": [
            "tkinter",
            "pymupdf",
            "fitz",
        ],
    },
}

# Setup definition
setup(
    name="PDF_CROP",
    version="1.0",
    description="Crop Pdf For Oncology",
    author="MD: Abdullah Al Mamun",
    author_email="pygemsbd@gmail.com",
    options=options,
    executables=executables,
)
