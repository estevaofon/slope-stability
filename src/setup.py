import sys
from cx_Freeze import setup, Executable

base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
        name = "simple_PyQt4",
        version = "0.1",
        description = "Sample cx_Freeze PyQt4 script",
        executables = [Executable("Estabilidade-de-Taludes.py", base = base)])
