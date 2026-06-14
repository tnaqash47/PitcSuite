import os
import sys

def config_path():
    """Always resolves config.ini next to the running EXE or script."""
    return os.path.join(os.path.dirname(sys.argv[0]), "config.ini")
