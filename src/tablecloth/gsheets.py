"""Write Google Sheets templates."""
from __future__ import annotations

import math
from contextlib import contextmanager
from typing import Any, Dict, Iterator, List, Literal

try:
    import pygsheets
    import pygsheets.client
    import pygsheets.exceptions
except ImportError:
    raise ImportError('Writing Google Sheets templates requires `pygsheets`')

from . import constants, helpers
from .layout import Layout

MAX_NAME_LENGTH: int = 100
"""Maximum length of sheet name."""

DEFAULT_SHEET_NAME: str = 'Sheet1'
"""Default sheet name."""


@contextmanager
def batched(client: pygsheets.client.Client) -> Iterator[pygsheets.client.Client]:
    """Context manager for updating Google Sheets in batch mode."""
    client.set_batch_mode(True)
    try:
        yield client
    except Exception as e:
        raise e
    else:
        client.run_batch()
    finally:
        client.set_batch_mode(False)


def calculate_minimum_cell_width(
    string: str,
    family: str = 'arial',
    size: float = 10,
    bold: bool = False,
    italic: bool = False,
    dpi: float = 96,
) -> float:
    """
    Calculate minimum cell width (pixels) to fit a string.

    Parameters
    ----------
    string
        String to measure.
    family
        Font family name (only 'arial' or 'calibri' is supported).
    size
        Font size in points.
    bold
        Whether cell is bold.
    italic
        Whether cell is italic.
    dpi
        Display dots per inch (dpi) assumed by Google Sheets (96 dpi).

    Raises
    ------
    NotImplementedError
        Font family is not supported.
    """
    # Load font widths
    family = family.lower()
    if family not in ['arial', 'calibri']:
        raise NotImplementedError(f'Font family {family} is not supported')
    font = family
    if bold:
        font = f'{font}-bold'
    if italic:
        font = f'{font}-italic'
    widths = constants.FONT_WIDTHS[font]
    # Measure only the longest line in a multiline string
    string = max(string.split('\n'), key=len)
    em = sum(widths[ord(char)] for char in string)
    points = size * em
    inches = points / 72
    # Google Sheets applies 3-pixel padding on each side
    return inches * dpi + 2 * 3


def calculate_column_width(header: str, **kwargs: Any) -> float:
    """
    Calculate column width (in pixels) from header.

    Width is the minimum width to fit the header plus 10 pixels padding,
    and no less than 70 pixels.

    Parameters
    ----------
    header
        Column header.
    **kwargs
        Additional keyword arguments for :func:`calculate_minimum_cell_width`.
    """
    minimum = calculate_minimum_cell_width(header, **kwargs)
    return max(minimum + 10, 70)


def write_table(
    sheet: pygsheets.Worksheet,
    header: List[str],
    freeze_header: bool = False,
    format_header: dict | None = None,
    header_height: int | float | None = None,
    comment_header: List[str] | None = None,
    hide_columns: bool = False,
    column_widths: List[int | float | None] | None = None,
) -> None:
    """
    Write an empty table (with header) to a Google Sheets sheet.

    Parameters
    ----------
    sheet
        Worksheet.
    header
        Content of the header row (one value per cell).
    freeze_header
        Whether to freeze the header.
    format_header
        Format to apply to header cells. See
        https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/cells?hl=en#cellformat
    header_height
        Header row height in pixels (default adjusts to content).
    comment_header
        Whether and what text to add as a note to each header cell (or None to skip).
    hide_columns
        Whether to hide unused columns.
    column_widths
        Width of each column in pixels (or None to leave unchanged).
        Default is the minimum width to fit the header plus 10 pixels padding,
        and no less than 70 pixels.
        See :func:`calculate_minimum_cell_width` as a starting point for customization.
    """
    ncols = len(header)
    header_range = pygsheets.DataRange((1, 1), (1, ncols), sheet)
    if not column_widths:
        # Load any existing cell text format to calculate column widths
        header_range.fetch()
    with batched(sheet.client):
        header_range.update_values([header])
        if format_header:
            header_range.apply_format(
                cell=None,
                fields='userEnteredFormat',
                cell_json={'userEnteredFormat': format_header},
            )
        if header_height is not None:
            sheet.adjust_row_height(1, 1, pixel_size=header_height)
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
        # Resize columns
        for i, content in enumerate(header, start=1):
            if column_widths:
                width = column_widths[i - 1]
            if not column_widths or width is None:
                # Determine final cell format
                format = header_range.cells[0][i - 1].text_format or {}
                if format_header:
                    format = {**format, **format_header.get('textFormat', {})}
                width = calculate_column_width(
                    content,
                    family=format.get('fontFamily', 'arial'),
                    size=format.get('fontSize', 10),
                    bold=format.get('bold', False),
                    italic=format.get('italic', False),
                )
            sheet.adjust_column_width(i, i, pixel_size=math.ceil(width))


def write_enum(sheet: pygsheets.Worksheet, values: list, col: int) -> None:
    """
    Write enumerated values to a column in a Google Sheets sheet.

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
    header_comments: Dict[str, List[str]] | None = None,
    dropdowns: bool = True,
    error_type: Literal['warning', 'stop'] | None = None,
    validate_foreign_keys: bool = True,
    format_invalid: dict
    | None = {
        'backgroundColorStyle': {'rgbColor': {'red': 1, 'green': 0.8, 'blue': 0.8}}
    },
    format_header: dict
    | None = {
        'textFormat': {'bold': True},
        'backgroundColorStyle': {'rgbColor': {'red': 0.8, 'green': 0.8, 'blue': 0.8}},
        'verticalAlignment': 'TOP',
    },
    header_height: int | float | None = None,
    freeze_header: bool = True,
    hide_columns: bool = False,
    column_widths: Dict[str, List[int | float | None]] | None = None,
) -> None:
    """
    Write a template for data entry to a Google Sheets workbook.

    A sheet with the default name 'Sheet1' is deleted if it is empty.

    Parameters
    ----------
    package
        Frictionless Data Tabular Data Package specification.
        See https://specs.frictionlessdata.io/tabular-data-package.
        Table names (``resource.name``) are used as sheet names and can be any string up
        to 100 characters long.
    book
        Workbook.
    enum_sheet
        Name of sheet used for enumerated value constraints.
    header_comments
        For each table (with `resource.name` as dictionary key),
        whether and what text to add as a note to each column header
        (one value per cell, None to skip).
    dropdowns
        Whether to display a dropdown if a column meets certain conditions.
        See :meth:`.Layout.select_column_dropdown`.
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
    header_height
        Header row height in pixels (default adjusts to content).
    freeze_header
        Whether to freeze the header.
    hide_columns
        Whether to hide unused columns.
    column_widths
        For each table (with `resource.name` as dictionary key),
        the widths in pixels of each column (one value per column, None to skip).
        Default is the minimum width to fit the header plus 10 pixels padding,
        and no less than 70 pixels.
        See :func:`calculate_minimum_cell_width` as a starting point for customization
    """
    # ---- Initialize
    layout = Layout.from_package(
        package=package,
        enum_sheet=enum_sheet,
        max_name_length=MAX_NAME_LENGTH,
        max_rows=None,
    )

    # ---- Write tables
    for table_props in layout.tables:
        sheet: pygsheets.Worksheet = book.add_worksheet(table_props['sheet'])
        write_table(
            sheet=sheet,
            header=table_props['columns'],
            comment_header=(header_comments or {}).get(table_props['table']),
            format_header=format_header,
            header_height=header_height,
            freeze_header=freeze_header,
            hide_columns=hide_columns,
            column_widths=(column_widths or {}).get(table_props['table']),
        )

    # --- Write enums
    if (dropdowns or error_type or format_invalid) and layout.enums:
        sheet = book.add_worksheet(layout.enum_sheet)
        sheet.hidden = True
        for enum_props in layout.enums:
            write_enum(sheet=sheet, values=enum_props['values'], col=enum_props['col'])

    # --- Add column checks
    with batched(book.client):
        for resource in package['resources']:
            table = resource['name']
            sheet_name = layout.get_table(table)['sheet']
            sheet = book.worksheet_by_title(sheet_name)
            foreign_keys = resource['schema'].get('foreignKeys', [])

            # For each column
            for field in resource['schema']['fields']:
                column = field['name']
                # HACK: Force pygsheets to accept one-sided unbounded range
                code = layout.get_column_code(table, column)
                cells = pygsheets.GridRange(start=code, end=code, worksheet=sheet)
                cells.set_json({'startRowIndex': 1, **cells.to_json()})
                dtype = field.get('type', 'any')
                constraints = field.get('constraints', {})

                # Data validation
                validation = None
                # Dropdown
                if dropdowns:
                    dropdown = layout.select_column_dropdown(
                        table=table,
                        column=column,
                        dtype=dtype,
                        constraints=constraints,
                        foreign_keys=foreign_keys,
                        indirect=False,
                    )
                    if dropdown:
                        range_validation = dropdown['source'] in ('enum', 'foreign_key')
                        validation = {
                            'condition_type': (
                                'ONE_OF_RANGE' if range_validation else 'ONE_OF_LIST'
                            ),
                            'condition_values': (
                                [f"={dropdown['options']}"]
                                if range_validation
                                else dropdown['options']
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
                        constraints=constraints,
                        foreign_keys=foreign_keys if validate_foreign_keys else None,
                        indirect=False,
                    )
                    check = helpers.build_column_validation(checks)
                    if check:
                        validation = {
                            'condition_type': 'CUSTOM_FORMULA',
                            'condition_values': [f"={check['formula']}"],
                            'strict': error_type == 'stop',
                            'showCustomUi': False,
                            'inputMessage': check['message'],
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
                        constraints=constraints,
                        foreign_keys=foreign_keys if validate_foreign_keys else None,
                        indirect=True,
                    )
                    if checks:
                        formula = helpers.build_column_condition(
                            checks=checks,
                            valid=False,
                            col=layout.get_column_code(table, column),
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
