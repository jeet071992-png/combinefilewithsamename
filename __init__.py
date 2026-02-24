"""
combinefilewithsamename
========================
Combine multiple Excel workbooks with a popup file selector.
Inject VBA macro into Excel automatically.

Usage:
    python -m combinefilewithsamename
    or
    from combinefilewithsamename import run
    run()
"""

from .core import run, inject_vba, combine_excel_files

__version__ = "1.0.0"
__author__ = "jeet071992-png"
__all__ = ["run", "inject_vba", "combine_excel_files"]
