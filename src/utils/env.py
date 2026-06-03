import sys
import os

def setup_environment():
    # 1. Prevent UnicodeEncodeError when printing UTF-8 characters to standard output in Windows console
    if sys.stdout is not None:
        try:
            sys.stdout.reconfigure(encoding='utf-8', errors='backslashreplace')
        except AttributeError:
            pass
    if sys.stderr is not None:
        try:
            sys.stderr.reconfigure(encoding='utf-8', errors='backslashreplace')
        except AttributeError:
            pass

    # 2. Automatically configure Tcl/Tk environment paths for Windows Python interpreters
    laragon_tcl = r"C:\laragon\bin\python\python-3.13\tcl\tcl8.6"
    laragon_tk = r"C:\laragon\bin\python\python-3.13\tcl\tk8.6"

    if os.path.exists(laragon_tcl):
        os.environ["TCL_LIBRARY"] = laragon_tcl
    if os.path.exists(laragon_tk):
        os.environ["TK_LIBRARY"] = laragon_tk
