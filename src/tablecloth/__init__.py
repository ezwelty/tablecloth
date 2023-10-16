"""Create spreadsheet templates for tabular data entry."""
from importlib.metadata import version as _version

from .layout import Layout

__all__ = ['Layout']
__version__ = _version(__name__)
