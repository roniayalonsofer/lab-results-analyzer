"""
code_reload.py
--------------
After a push, Streamlit Cloud re-runs the new app.py in the same process, while
modules it already imported stay in memory at their old version. Dropping them
when the code on disk changes makes the next import load the new code.
"""

import importlib
import os
import sys


def drop_stale_modules(root: str) -> float:
    """Forget modules under *root* if its code changed. Returns the code's version."""
    root = os.path.abspath(root)
    version = max(
        (os.path.getmtime(os.path.join(folder, name))
         for folder, _, names in os.walk(root) for name in names if name.endswith(".py")),
        default=0.0,
    )
    if getattr(sys, "_lab_code_version", None) != version:
        for name, module in list(sys.modules.items()):
            path = getattr(module, "__file__", None)
            if path and os.path.abspath(path).startswith(root + os.sep):
                del sys.modules[name]
        importlib.invalidate_caches()
        sys._lab_code_version = version
    return version
