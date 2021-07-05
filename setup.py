import sys
from cx_Freeze import setup, Executable
import os

# <added>
import os.path
PYTHON_INSTALL_DIR = os.path.dirname(os.path.dirname(os.__file__))
os.environ['TCL_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tcl8.6')
os.environ['TK_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tk8.6')
# </added>

base = None
if sys.platform == 'win32':
    base = 'Win32GUI'

executables = [
    Executable('PyGems Excel Merger.py', shortcutName="PyGems Excel Merger",
            icon="dicon.ico",base=base)
]


# Now create the table dictionary
shortcut_table = [
    ("DesktopShortcut",        # Shortcut
     "DesktopFolder",          # Directory_
     "PyGems Excel Merger",           # Name
     "TARGETDIR",              # Component_
     "[TARGETDIR]PyGems Excel Merger.exe",# Target
     None,                     # Arguments
     None,                     # Description
     None,                     # Hotkey
     None,                     # Icon
     None,                     # IconIndex
     None,                     # ShowCmd
     'TARGETDIR'               # WkDir
     )
    ]

# Now create the table dictionary
msi_data = {"Shortcut": shortcut_table}
bdist_msi_options = {'data': msi_data}

# <added>
options = {
    
          "bdist_msi": bdist_msi_options,

    'build_exe': {
        'include_files':[
            os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tk86t.dll'),
            os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tcl86t.dll'),
	    'question.png','dicon.ico','software.png','xlsx.png','tips.png',

         ],
         
        
    },
}
# </added>


setup(name = 'PyGems Excel Merger',
      version = "1.0",
      description = 'PyGems Excel Merger is a simple excel merging tool that built with pure python language.Copyright © 2019-2020',
      # <added>
      author = 'MD : ABDULLAH AL MAMUN',
      author_email = 'pygemsbd@gmail.com',
      options = options,
      # </added>
      executables = executables
      )
