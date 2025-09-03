"""PyInstaller launcher for PST Splitter GUI.

This wrapper ensures the package 'pstsplitter' is imported as a package so that
its internal modules using relative imports (e.g. from .util import ...) work
in the frozen executable. Directly freezing 'pstsplitter/gui.py' caused
'ImportError: attempted relative import with no known parent package' because
PyInstaller executed the module as a top-level script (no __package__).
"""
from pstsplitter.gui import run_app

if __name__ == "__main__":  # pragma: no cover
    run_app()
