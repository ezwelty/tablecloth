"""Test cases for the gsheets module."""
import os
from pathlib import Path

import dotenv
import pygsheets
import pygsheets.exceptions
import pytest
import yaml

import tablecloth.gsheets

dotenv.load_dotenv()


if os.getenv('GOOGLE_CLOUD_SERVICE_ACCOUNT_KEY'):
    try:
        client = pygsheets.authorize(
            service_account_env_var='GOOGLE_CLOUD_SERVICE_ACCOUNT_KEY'
        )
    except Exception as e:
        pytest.skip(
            f'Google Cloud authorization failed: {e}',
            allow_module_level=True,
        )
else:
    pytest.skip(
        'GOOGLE_CLOUD_SERVICE_ACCOUNT_KEY environment variable is empty',
        allow_module_level=True,
    )


def get_book(name: str) -> pygsheets.Spreadsheet:
    """Open or create a Google Sheets workbook by name."""
    try:
        # Open and reset if already exists
        book = client.open(name)
        tablecloth.gsheets.reset_sheets(book)
    except pygsheets.exceptions.SpreadsheetNotFound:
        # Create and share if not
        book = client.create(name)
    # Share with user's Google Account
    share_with = os.getenv('GOOGLE_ACCOUNT')
    if share_with:
        book.share(share_with, role='writer', type='user')
    # Print full URL to spreadsheet
    print(f'https://docs.google.com/spreadsheets/d/{book.id}')
    return book


@pytest.mark.gsheets
@pytest.mark.parametrize(
    'name, arguments',
    [('warning-hide-columns', {'hide_columns': True, 'error_type': 'warning'})],
)
def test_writes_template(name: str, arguments: dict) -> None:
    """It writes a Google Sheets template."""
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
    # Write template
    book = get_book(f'tablecloth-test-{name}')
    kwargs = {'header_comments': comments, **arguments}
    tablecloth.gsheets.write_template(package, book=book, **kwargs)
