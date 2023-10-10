from collections import defaultdict
from contextlib import contextmanager
from typing import Dict, List, Literal, Optional

try:
    import pygsheets
    import pygsheets.client
except ImportError:
    raise ImportError('Writing Google Sheets templates requires `pygsheets`')

from . import constants, helpers
from .layout import Layout

MAX_NAME_LENGTH: int = 100
"""Maximum length of sheet name."""

DEFAULT_SHEET_NAME: str = 'Sheet1'
"""Default sheet name."""


@contextmanager
def batched(client: pygsheets.client.Client) -> None:
    """
    Context manager for updating Google Sheets in batch mode.
    """
    client.set_batch_mode(True)
    try:
        yield client
    except Exception as e:
        raise e
    else:
        client.run_batch()
    finally:
        client.set_batch_mode(False)


def write_table(
    sheet: pygsheets.Worksheet,
    header: List[str],
    freeze_header: bool = False,
    format_header: dict = None,
    comment_header: List[str] = None,
    hide_columns: bool = False,
) -> None:
    """
    Write empty table with header to a Google Sheets sheet.

    Parameters
    ----------
    sheet
        Worksheet.
    header
        Content of the header row (one value per cell).
    freeze_header
        Whether to freeze the header.
    format_header
        Whether and how to format header cells. See the Google API reference:
        https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/cells?hl=en#cellformat
    comment_header
        Whether and what text to add as a note to each header cell
        (one value per cell, None to skip).
    hide_columns
        Whether to hide unused columns.
    """
    ncols = len(header)
    header_range = pygsheets.DataRange((1, 1), (1, ncols), sheet)
    with batched(sheet.client):
        header_range.update_values([header])
        if format_header:
            header_range.apply_format(
                cell=None,
                fields='userEnteredFormat',
                cell_json={'userEnteredFormat': format_header},
            )
        if freeze_header:
            sheet.frozen_rows = 1
        if comment_header:
            for i, note in enumerate(comment_header, start=1):
                if note:
                    pygsheets.DataRange((1, i), (1, i), sheet).apply_format(
                        cell=None, fields='note', cell_json={'note': note}
                    )
        if hide_columns:
            sheet.resize(cols=ncols)
        # Resize column widths to fit header values with some padding
        # (automatic widths results in very narrow columns for small values)
        for i, value in enumerate(header, start=1):
            width = int(max(10, len(value) * 1.2) * 7.7)
            sheet.adjust_column_width(i, i, pixel_size=width)


def write_enum(sheet: pygsheets.Worksheet, values: list, col: int) -> None:
    """
    Write enumerated values to a Google Sheets sheet.

    Parameters
    ----------
    sheet
        Worksheet.
    values
        Values.
    col
        Column to write values to (zero-indexed).
    """
    nrows = len(values)
    enum_range = pygsheets.DataRange((1, col + 1), (nrows, col + 1), sheet)
    enum_range.update_values([[value] for value in values])


def reset_sheets(book: pygsheets.Spreadsheet) -> None:
    """
    Reset a Google Sheets workbook.

    Deletes all sheets and replaces them with a blank sheet with name 'Sheet1'.

    Parameters
    ----------
    book
        Workbook.
    """
    sheets = book.worksheets()
    # Delete hidden sheets first
    sheets = sorted(sheets, key=lambda x: not x.hidden)
    with batched(book.client):
        for sheet in sheets[:-1]:
            book.del_worksheet(sheet)
    # To delete the last sheet: add a new blank sheet, delete the old one, then rename
    old_sheet = sheets[-1]
    new_sheet = book.add_worksheet(f'_{old_sheet.title}')
    book.del_worksheet(old_sheet)
    new_sheet.title = DEFAULT_SHEET_NAME


def write_template(
    package: dict,
    book: pygsheets.Spreadsheet,
    enum_sheet: str = 'lists',
    header_comments: Dict[str, List[str]] = None,
    dropdowns: bool = True,
    error_type: Literal['warning', 'stop'] = None,
    validate_foreign_keys: bool = True,
    format_invalid: Optional[dict] = {
        'backgroundColorStyle': {'rgbColor': {'red': 1, 'green': 0.8, 'blue': 0.8}}
    },
    format_header: Optional[dict] = {
        'textFormat': {'bold': True},
        'backgroundColorStyle': {'rgbColor': {'red': 0.8, 'green': 0.8, 'blue': 0.8}},
    },
    freeze_header: bool = True,
    hide_columns: bool = False,
) -> None:
    """
    Write package template to a Google Sheets workbook.

    A sheet with the default name 'Sheet1' is deleted if it is empty.

    Parameters
    ----------
    package
        Frictionless Data Tabular Data Package specification.
        See https://specs.frictionlessdata.io/tabular-data-package.
    book
        Workbook.
    enum_sheet
        Name of sheet used for enumerated value constraints.
    header_comments
        For each table (with `resource.name` as dictionary key),
        whether and what text to add as a note to each column header
        (one value per cell, None to skip).
    dropdowns
        Whether to display a dropdown if a column meets one of the following conditions:

        * Boolean data type (`field.type: 'boolean'`). Shown as (TRUE, FALSE).
        * Enumerated value constraint (`field.constraints.enum`).
        * Foreign key (`resource.schema.foreignKeys`). Foreign keys involving multiple
          columns are treated as multiple single-column foreign keys. If a column has
          multiple foreign keys, only the first is used.

        If selected, additional column constraints will not be enforced by `error_type`,
        but will be by `format_invalid`.
    error_type
        Whether and which error type to raise if invalid input is detected:

        * warning: Leave cell value unchanged but display a (generic) warning on hover.
        * stop: Reset cell value and display a (custom) error.

        If `error_type` is `None` but `dropdowns` is True, dropdowns will display
        warnings because Google Sheets dropdowns require an error type.
    validate_foreign_keys
        Whether to validate foreign keys (True) or only use for dropdowns (False).
        In the latter case, a warning will still be shown for a value not in the list
        (see `error_type`).
    format_invalid
        Formatting for cells with invalid input. See
        https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/cells#CellFormat.
    format_header
        Formatting for header cells. See
        https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/cells#CellFormat.
    freeze_header
        Whether to freeze the header.
    hide_columns
        Whether to hide unused columns.
    """
    # ---- Initialize
    layout = Layout(
        enum_sheet=enum_sheet, max_name_length=MAX_NAME_LENGTH, max_rows=None
    )

    # ---- Write tables
    for resource in package['resources']:
        table = resource['name']
        columns = [field['name'] for field in resource['schema']['fields']]
        layout.set_table(table, columns)
        sheet = book.add_worksheet(table)
        write_table(
            sheet,
            header=[field['name'] for field in resource['schema']['fields']],
            freeze_header=freeze_header,
            format_header=format_header,
            comment_header=header_comments.get(table) if header_comments else None,
            hide_columns=hide_columns,
        )

    # --- Write enums
    # Add separately so that the remaining can be batched
    if dropdowns or error_type or format_invalid:
        for resource in package['resources']:
            for field in resource['schema']['fields']:
                enum = field.get('constraints', {}).get('enum')
                if enum:
                    layout.set_enum(enum)
                    try:
                        esheet = book.worksheet_by_title(enum_sheet)
                    except pygsheets.exceptions.WorksheetNotFound:
                        esheet = book.add_worksheet(enum_sheet)
                        esheet.hidden = True
                    write_enum(
                        sheet=esheet, values=enum, col=layout.get_enum(enum)['col']
                    )

    # --- Add column checks
    with batched(book.client):
        for resource in package['resources']:
            table = resource['name']
            sheet: pygsheets.Worksheet = book.worksheet_by_title(table)

            # Compile foreign keys by column
            foreign_keys = defaultdict(list)
            for key in resource['schema'].get('foreignKeys', []):
                fields = helpers.to_list(key['fields'])
                ref_fields = helpers.to_list(key['reference']['fields'])
                # Composite keys are treated as multiple simple keys
                for field, ref_field in zip(fields, ref_fields):
                    foreign_keys[field].append(
                        (key['reference']['resource'] or table, ref_field)
                    )

            # For each column
            for field in resource['schema']['fields']:
                column = field['name']
                # HACK: Force pygsheets to accept one-sided unbounded range
                code = layout.get_column_code(table, column)
                cells = pygsheets.GridRange(start=code, end=code, worksheet=sheet)
                cells.set_json({'startRowIndex': 1, **cells.to_json()})
                dtype = field.get('type', 'any')
                enum = field.get('constraints', {}).get('enum')
                constraints = {
                    helpers.camel_to_snake_case(k): v
                    for k, v in field.get('constraints', {}).items()
                    if helpers.camel_to_snake_case(k) in constants.CONSTRAINTS
                }

                # Data validation
                validation = None
                range_validation = None
                # Dropdown
                if dropdowns:
                    values = None
                    if dtype == 'boolean':
                        values = ['TRUE', 'FALSE']
                    elif column in foreign_keys:
                        # TODO: Move to Layout.get_column_dropdown
                        # NOTE: Uses first foreign key
                        foreign_table, foreign_column = foreign_keys[column][0]
                        values = layout.get_column_range(
                            foreign_table,
                            foreign_column,
                            fixed=True,
                            absolute=True,
                            indirect=False,
                        )
                        range_validation = True
                    elif enum:
                        values = layout.get_enum_range(enum, indirect=False)
                        range_validation = True
                    if values:
                        validation = {
                            'condition_type': (
                                'ONE_OF_RANGE' if range_validation else 'ONE_OF_LIST'
                            ),
                            'condition_values': (
                                [f'={values}'] if range_validation else values
                            ),
                            'strict': error_type == 'stop',
                            'showCustomUi': True,
                            'inputMessage': 'Value must be in the dropdown list',
                        }
                if not validation and error_type:
                    # Get column checks
                    checks = layout.gather_column_checks(
                        table,
                        column,
                        valid=True,
                        dtype=dtype,
                        **constraints,
                        enum=enum,
                        foreign_keys=(
                            foreign_keys.get(column) if validate_foreign_keys else None
                        ),
                        indirect=False,
                    )
                    validation = helpers.build_column_validation(checks)
                    if validation:
                        validation = {
                            'condition_type': 'CUSTOM_FORMULA',
                            'condition_values': [f"={validation['formula']}"],
                            'strict': error_type == 'stop',
                            'showCustomUi': False,
                            'inputMessage': validation['message'],
                        }
                if validation:
                    sheet.set_data_validation(grange=cells, **validation)

                # Conditional formatting
                if format_invalid:
                    checks = layout.gather_column_checks(
                        table,
                        column,
                        valid=False,
                        dtype=dtype,
                        **constraints,
                        enum=enum,
                        foreign_keys=(
                            foreign_keys.get(column) if validate_foreign_keys else None
                        ),
                        indirect=True,
                    )
                    formula = helpers.build_column_condition(checks, valid=False)
                    if formula:
                        formula = formula.format(
                            col=layout.get_column_code(table, column), row=2
                        )
                        sheet.add_conditional_formatting(
                            # HACK: Force pygsheets to use grange
                            start=None,
                            end=None,
                            grange=cells,
                            format=format_invalid,
                            condition_type='CUSTOM_FORMULA',
                            condition_values=[f'={formula}'],
                        )

    # ---- Delete default sheet
    # if present, not used by package, and empty
    try:
        sheet = book.worksheet_by_title(DEFAULT_SHEET_NAME)
    except pygsheets.exceptions.WorksheetNotFound:
        return
    else:
        resource_names = [resource['name'] for resource in package['resources']]
        if (
            resource_names
            and DEFAULT_SHEET_NAME not in resource_names
            and DEFAULT_SHEET_NAME != enum_sheet
            and sheet.get_all_values(
                include_tailing_empty_rows=False, include_tailing_empty=False
            )
            == [[]]
        ):
            book.del_worksheet(sheet)
