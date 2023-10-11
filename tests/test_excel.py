from __future__ import annotations

import datetime
import zipfile
from pathlib import Path
from typing import Any, List

import pytest
import xlsxwriter
import yaml

import tablecloth.excel


def write_indempotent_template(path: str | Path, **kwargs: Any) -> None:
    # Load package descriptor
    package_path = Path(__file__).parent / 'datapackage.yaml'
    package = yaml.safe_load(package_path.read_text(encoding='utf-8'))
    # Build header comments
    comments = {
        resource['name']: [
            field['description'] for field in resource['schema']['fields']
        ]
        for resource in package['resources']
    }
    book: xlsxwriter.Workbook = tablecloth.excel.write_template(
        package, header_comments=comments, **kwargs
    )
    book.set_properties({'created': datetime.datetime(2000, 1, 1)})
    book.filename = path
    book.close()


def read_xlsx_as_string(path: str | Path) -> str:
    def read_children(root: zipfile.Path) -> str:
        text = ''
        inner_paths: List[zipfile.Path] = sorted(root.iterdir(), key=lambda x: x.name)
        for inner_path in inner_paths:
            if inner_path.is_dir():
                text += read_children(inner_path)
            else:
                text += f'##### {inner_path.name} #####\n\n'
                # Cheap pretty print
                text += inner_path.read_text(encoding='utf-8').replace('><', '>\n<')
                text += '\n\n'
        return text

    root = zipfile.Path(path)
    return read_children(root)


@pytest.mark.parametrize(
    'name, arguments',
    [
        ('default', {}),
        ('errors', {'error_type': 'information'}),
        (
            'errors-except-fkeys',
            {'error_type': 'information', 'validate_foreign_keys': False},
        ),
    ],
)
def test_writes_template(name: str, arguments: dict) -> None:
    expected = Path(__file__).parent / 'xlsx' / f'{name}.xlsx'
    actual = Path(__file__).parent / 'xlsx' / f'{name}-test.xlsx'
    write_indempotent_template(path=actual, **arguments)
    assert read_xlsx_as_string(actual) == read_xlsx_as_string(expected)
