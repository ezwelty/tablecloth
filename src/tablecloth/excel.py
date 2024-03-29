"""Write Microsoft Excel templates."""
from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Literal, cast

try:
    import xlsxwriter
    import xlsxwriter.format
    import xlsxwriter.worksheet
except ImportError:
    raise ImportError('Writing Excel templates requires `xlsxwriter`')

from . import constants, helpers
from .layout import Layout

MAX_COLS: int = 16384
"""Maximum number of columns per sheet."""

MAX_ROWS: int = 1048576
"""Maximum number of rows per sheet."""

MAX_NAME_LENGTH: int = 31
"""Maximum length of sheet name."""


def write_table(
    sheet: xlsxwriter.worksheet.Worksheet,
    header: List[str],
    freeze_header: bool = False,
    format_header: xlsxwriter.format.Format | None = None,
    comment_header: List[str] | None = None,
    format_comments: dict | None = None,
    hide_columns: bool = False,
) -> None:
    """
    Write an empty table (with header) to a Microsoft Excel sheet.

    Parameters
    ----------
    sheet
        Worksheet.
    header
        Content of the header row (one value per cell).
    freeze_header
        Whether to freeze the header.
    format_header
        Whether and how to format header cells.
        See https://xlsxwriter.readthedocs.io/format.html.
    comment_header
        Whether and what text to add as a comment to each header cell (or None to skip).
    format_comments
        Whether and how to format header comments. See
        https://xlsxwriter.readthedocs.io/working_with_cell_comments.html#cell-comments.
    hide_columns
        Whether to hide unused columns.
        Ignored if `comment_header` is True, as Excel prevents hiding columns that
        overlap a comment.
    """
    # Write header
    sheet.write_row(0, 0, header, format_header)
    # Hide unused columns
    if hide_columns and not comment_header:
        sheet.set_column(len(header), MAX_COLS - 1, options={'hidden': 1})
    # Freeze header
    if freeze_header:
        sheet.freeze_panes(1, 0)
    # Resize columns and add header comments
    for i, content in enumerate(header):
        width = max(10, len(content) * 1.2)
        sheet.set_column(i, i, width=width)
        if comment_header and comment_header[i]:
            sheet.write_comment(0, i, comment_header[i], format_comments)


def write_enum(sheet: xlsxwriter.worksheet.Worksheet, values: list, col: int) -> None:
    """
    Write enumerated values to a column in a Microsoft Excel sheet.

    Parameters
    ----------
    sheet
        Worksheet.
    values
        Values.
    col
        Column to write values to (zero-indexed).
    """
    for i, value in enumerate(values):
        if isinstance(value, str):
            sheet.write_string(i, col, value)
        else:
            sheet.write(i, col, value)


def write_template(
    package: dict,
    path: str | Path | None = None,
    enum_sheet: str = 'lists',
    header_comments: Dict[str, List[str]] | None = None,
    dropdowns: bool = True,
    error_type: Literal['information', 'warning', 'stop'] | None = None,
    validate_foreign_keys: bool = True,
    format_invalid: dict | None = {'bg_color': '#ffc7ce'},
    format_header: dict | None = {'bold': True, 'bg_color': '#d3d3d3'},
    format_comments: dict | None = {'font_size': 11, 'x_scale': 2, 'y_scale': 2},
    freeze_header: bool = True,
    hide_columns: bool = False,
) -> xlsxwriter.Workbook | None:
    """
    Write a template for data entry to a Microsoft Excel workbook.

    Parameters
    ----------
    package
        Frictionless Data Tabular Data Package specification.
        See https://specs.frictionlessdata.io/tabular-data-package.
        Table names (``resource.name``) are used as sheet names and can be any string up
        to 31 characters long.
    path
        Path of the created Microsoft Excel (.xlsx) file.
        If None, the result is returned as `xlsxwriter.Workbook`.
    enum_sheet
        Name of sheet used for enumerated value constraints.
    header_comments
        For each table (with `resource.name` as dictionary key),
        whether and what text to add as a comment to each column header
        (one value per cell, None to skip).
    dropdowns
        Whether to display a dropdown if a column meets certain conditions.
        See :meth:`.Layout.select_column_dropdown`.
        If selected, additional column constraints will not be enforced by `error_type`,
        but will be by `format_invalid`.
    error_type
        Whether and which error type to raise if invalid input is detected:

        * information: Display a message with a button to accept.
        * warning: Display a message with buttons to accept or retry.
        * stop: Display a message with buttons to cancel or retry.
    validate_foreign_keys
        Whether to validate foreign keys (True) or only use for dropdowns (False).
    format_invalid
        Formatting for cells with invalid input.
        See https://xlsxwriter.readthedocs.io/format.html.
    format_header
        Formatting for header cells.
        See https://xlsxwriter.readthedocs.io/format.html.
    format_comments
        Whether and how to format header comments. See
        https://xlsxwriter.readthedocs.io/working_with_cell_comments.html#cell-comments.
    freeze_header
        Whether to freeze the header.
    hide_columns
        Whether to hide unused columns.
    """
    # ---- Initialize
    layout = Layout.from_package(
        package=package,
        enum_sheet=enum_sheet,
        max_name_length=MAX_NAME_LENGTH,
        max_rows=MAX_ROWS,
    )
    book = xlsxwriter.Workbook(filename=path)
    # Register formats
    if format_invalid:
        invalid_format = book.add_format(format_invalid)
    if format_header:
        header_format = book.add_format(format_header)

    # ---- Write tables
    for table_props in layout.tables:
        sheet: xlsxwriter.worksheet.Worksheet = book.add_worksheet(table_props['sheet'])
        write_table(
            sheet=sheet,
            header=table_props['columns'],
            comment_header=(header_comments or {}).get(table_props['table']),
            format_header=header_format,
            format_comments=format_comments,
            freeze_header=freeze_header,
            hide_columns=hide_columns,
        )

    # ---- Write enums
    if (dropdowns or error_type or format_invalid) and layout.enums:
        sheet = book.add_worksheet(layout.enum_sheet)
        sheet.hide()
        for enum_props in layout.enums:
            write_enum(sheet=sheet, values=enum_props['values'], col=enum_props['col'])

    # --- Add column checks
    for resource in package['resources']:
        table = resource['name']
        sheet_name = layout.get_table(table)['sheet']
        sheet = book.get_worksheet_by_name(sheet_name)
        foreign_keys = resource['schema'].get('foreignKeys', [])

        # For each column
        for field in resource['schema']['fields']:
            column = field['name']
            cells = layout.get_column_range(table, column)
            dtype = field.get('type', 'any')
            constraints = cast(
                constants.Constraints,
                {
                    key: value
                    for key, value in field.get('constraints', {}).items()
                    # No regex support in Excel
                    if key not in ['pattern']
                },
            )

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
                    validation = {
                        'validate': 'list',
                        'value': dropdown['options'],
                        'error_title': 'Invalid value',
                        'error_message': 'Value must be in the dropdown list',
                        'ignore_blank': True,
                        'error_type': error_type or 'information',
                        'show_error': (
                            bool(error_type)
                            and (
                                dropdown['source'] != 'foreign_key'
                                or validate_foreign_keys
                            )
                        ),
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
                )
                check = helpers.build_column_validation(checks)
                if check:
                    validation = {
                        'validate': 'custom',
                        'value': check['formula'],
                        'error_title': 'Invalid value',
                        'error_message': check['message'],
                        'ignore_blank': check['ignore_blank'],
                        'error_type': error_type,
                        'show_error': bool(error_type),
                    }
            if validation:
                sheet.data_validation(cells, validation)

            # Conditional formatting
            if format_invalid:
                checks = layout.gather_column_checks(
                    table,
                    column,
                    valid=False,
                    dtype=dtype,
                    constraints=constraints,
                    foreign_keys=foreign_keys if validate_foreign_keys else None,
                )
                if checks:
                    formula = helpers.build_column_condition(
                        checks=checks,
                        valid=False,
                        col=layout.get_column_code(table, column),
                    )
                    sheet.conditional_format(
                        cells,
                        options={
                            'type': 'formula',
                            'criteria': formula,
                            'format': invalid_format,
                        },
                    )
    if path is None:
        return book
    book.close()
    return None
