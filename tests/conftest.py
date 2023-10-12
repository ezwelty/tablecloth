"""Configure pytest."""
from typing import List

import pytest


def pytest_configure(config: pytest.Config) -> None:
    """Add `gsheets` marker to pytest."""
    config.addinivalue_line('markers', 'gsheets: Run Google Sheets tests')


def pytest_addoption(parser: pytest.Parser) -> None:
    """Add `--gsheets` option to pytest."""
    parser.addoption(
        '--gsheets', action='store_true', default=False, help='Run Google Sheets tests'
    )


def pytest_collection_modifyitems(
    config: pytest.Config, items: List[pytest.Item]
) -> None:
    """Skip Google Sheets tests if `--gsheets` option is not provided."""
    if config.getoption('--gsheets'):
        return
    skip_gsheets = pytest.mark.skip(reason='Need --gsheets option to run')
    for item in items:
        if 'gsheets' in item.keywords:
            item.add_marker(skip_gsheets)
