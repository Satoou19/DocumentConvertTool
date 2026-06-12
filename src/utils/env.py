import sys
import os

def setup_environment():
    # 1. Prevent UnicodeEncodeError when printing UTF-8 characters to standard output in Windows console
    if sys.stdout is not None:
        try:
            sys.stdout.reconfigure(encoding='utf-8', errors='backslashreplace')  # type: ignore
        except AttributeError:
            pass
    if sys.stderr is not None:
        try:
            sys.stderr.reconfigure(encoding='utf-8', errors='backslashreplace')  # type: ignore
        except AttributeError:
            pass

    # 2. Automatically configure Tcl/Tk environment paths for Windows Python interpreters
    tcl_base = os.path.join(sys.prefix, "tcl")
    if os.path.exists(tcl_base):
        for entry in os.listdir(tcl_base):
            full = os.path.join(tcl_base, entry)
            if entry.startswith("tcl") and os.path.isdir(full):
                os.environ["TCL_LIBRARY"] = full
            if entry.startswith("tk") and os.path.isdir(full):
                os.environ["TK_LIBRARY"] = full
