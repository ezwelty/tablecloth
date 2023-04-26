import datetime
from typing import Any, List, Union
from pathlib import Path
import zipfile

import pytest
import yaml

import tablecloth


def write_indempotent_template(path: Union[str, Path], **kwargs: Any) -> None:
    # Load package descriptor
    package_path = Path(__file__).parent / 'datapackage.yaml'
    package = yaml.safe_load(package_path.read_text())
    # Build header comments
    comments = {
        resource['name']: [
            field['description'] for field in resource['schema']['fields']
        ]
        for resource in package['resources']
    }
    book = tablecloth.excel.write_template(package, header_comments=comments, **kwargs)
    book.set_properties({'created': datetime.datetime(2000, 1, 1)})
    book.filename = path
    book.close()


def read_xlsx_as_string(path: Union[str, Path]) -> str:
    def read_children(root: zipfile.Path) -> str:
        text = ''
        paths: List[zipfile.Path] = sorted(root.iterdir(), key=lambda x: x.name)
        for path in paths:
            if path.is_dir():
                text += read_children(path)
            else:
                text += f'##### {path.at} #####\n\n'
                # Cheap pretty print
                text += path.read_text().replace('><', '>\n<')
                text += '\n\n'
        return text

    root = zipfile.Path(path)
    return read_children(root)


@pytest.mark.parametrize(
    'name, arguments',
    [
        ('default', {}),
        ('errors', {'error_type': 'information'}),
    ],
)
def test_writes_template(name: str, arguments: dict) -> None:
    expected = Path(__file__).parent / 'xlsx' / f'{name}.xlsx'
    actual = Path(__file__).parent / 'xlsx' / f'{name}-test.xlsx'
    write_indempotent_template(path=actual, **arguments)
    assert read_xlsx_as_string(actual) == read_xlsx_as_string(expected)
