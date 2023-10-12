"""
Tests of missing optional dependencies.

To avoid having to create different test environments, missing dependencies are mocked
by manipulating the `sys.modules` dictionary.
"""
import sys

import pytest


def test_fails_to_import_excel_module_if_dependencies_are_missing() -> None:
    """It fails to import the excel module if dependencies are missing."""
    module, dependency = None, None
    if 'tablecloth.excel' in sys.modules:
        module = sys.modules['tablecloth.excel']
        del sys.modules['tablecloth.excel']
    if 'xlsxwriter' in sys.modules:
        dependency = sys.modules['xlsxwriter']
    sys.modules['xlsxwriter'] = None  # type: ignore
    try:
        with pytest.raises(ImportError):
            import tablecloth.excel  # noqa: F401
    finally:
        if module is not None:
            sys.modules['tablecloth.excel'] = module
        if dependency is not None:
            sys.modules['xlsxwriter'] = dependency


def test_fails_to_import_gsheets_module_if_dependencies_are_missing() -> None:
    """It fails to import the gsheets module if dependencies are missing."""
    module, dependency = None, None
    if 'tablecloth.gsheets' in sys.modules:
        module = sys.modules['tablecloth.gsheets']
        del sys.modules['tablecloth.gsheets']
    if 'pygsheets' in sys.modules:
        dependency = sys.modules['pygsheets']
    sys.modules['pygsheets'] = None  # type: ignore
    try:
        with pytest.raises(ImportError):
            import tablecloth.gsheets  # noqa: F401
    finally:
        if module is not None:
            sys.modules['tablecloth.gsheets'] = module
        if dependency is not None:
            sys.modules['pygsheets'] = dependency
