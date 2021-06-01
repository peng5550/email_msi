import sys
import os.path
from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need fine tuning.
PYTHON_INSTALL_DIR = os.path.dirname(os.path.dirname(os.__file__))
os.environ['TCL_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tcl8.6')
os.environ['TK_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tk8.6')

# Dependencies are automatically detected, but it might need fine tuning.
build_exe_options = {
    "packages": ["tkinter", "mttkinter", "threading", "openpyxl", "multiprocessing", "smtplib", "email"],
    "includes": ["tkinter"],
    'include_files': [
        os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tcl86t.dll'),
        os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tk86t.dll'),
        # os.path.join(os.path.dirname(__file__), "emailtools.py"),
        # os.path.join(os.path.dirname(__file__), "configRead.py"),
        # os.path.join(os.path.dirname(__file__), "config.ini"),
        os.path.join(os.path.dirname(__file__), "template"),
        os.path.join(os.path.dirname(__file__), "template_info.xlsx"),
        os.path.join(os.path.dirname(__file__), "demo.xlsx"),

    ]
    }

# GUI applications require a different base on Windows (the default is for a
# console application)
base = None
if sys.platform == "win32":
    base = "Win32GUI"

# "bdist_msi": bdist_msi_options
setup(name="EmailTools 1.0",
      version="1.0",
      description="Email Sender",
      options={"build_exe": build_exe_options},
      executables=[Executable("app.py",
                              shortcutName="EmailTools 1.0",
                              shortcutDir="DesktopFolder",
                              base=base)])
