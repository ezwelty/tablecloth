"""Write Microsoft Excel templates."""
from __future__ import annotations

import math
from pathlib import Path
from typing import Any, Dict, List, Literal, cast

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


def calculate_minimum_cell_width(
    string: str,
    family: str = 'calibri',
    size: float = 11,
    bold: bool = False,
    italic: bool = False,
) -> float:
    """
    Calculate minimum cell width (in character units) to fit a string.

    Calculation of string width in pixels seems exact, but the calculation of padding
    and conversion to character units is approximate and may not work in all Excel
    versions, screen resolutions, etc.

    Assumes that the user's default font is Calibri, 11 point, non-bold, non-italic
    (this is the default in Excel 2007+).

    Parameters
    ----------
    string
        String to measure.
    family
        Font family name. See :data:`.constants.FONT_FAMILIES` for which are supported.
    size
        Font size in points.
    bold
        Whether cell is bold.
    italic
        Whether cell is italic.

    Raises
    ------
    NotImplementedError
        Font family is not supported.
    """
    # DPI arbitrary after conversion to character units, but included for clarity
    dpi = 96
    # Assume default font is Calibri, 11 point, non-bold, non-italic
    default_font = 'calibri'
    default_size = 11
    # Load font widths
    font = family.lower()
    if font not in constants.FONT_FAMILIES:
        raise NotImplementedError(
            f"Font family '{family}' is not supported. "
            f'Use one of {constants.FONT_FAMILIES}'
        )
    if bold:
        font = f'{font}-bold'
    if italic:
        font = f'{font}-italic'
    widths = constants.FONT_WIDTHS[font]
    # Measure only the longest line in a multiline string
    string = max(string.split('\n'), key=len)
    em = sum(widths[ord(char)] for char in string)
    content_px = round(em * size / 72 * dpi)
    # Compute padding relative to '0' character width
    # https://stackoverflow.com/a/61041831
    zero_px = widths[ord('0')] * size / 72 * dpi
    pad_px = round((zero_px + 1) / 4) * 2 + 1
    px = content_px + pad_px
    # Convert to character units based on '0' character width in default font size
    default_widths = constants.FONT_WIDTHS[default_font]
    default_zero_px = default_widths[ord('0')] * default_size / 72 * dpi
    default_pad_px = round((default_zero_px + 1) / 4) * 2 + 1
    default_zero_pad_px = zero_px + pad_px
    return (
        px / math.ceil(default_zero_pad_px)
        if px < default_zero_pad_px
        else (px - default_pad_px) / math.ceil(default_zero_px)
    )


def calculate_column_width(header: str, wrap: bool = False, **kwargs: Any) -> float:
    """
    Calculate column width (in character units) from header.

    Width is the minimum width to fit the header plus 1.5 character units padding,
    and no less than 9 character units.

    Parameters
    ----------
    header
        Column header.
    wrap
        Whether text wrapping is enabled.
    **kwargs
        Additional keyword arguments for :func:`calculate_minimum_cell_width`.
    """
    if not wrap:
        header = header.replace('\n', '')
    minimum = calculate_minimum_cell_width(header, **kwargs)
    return max(minimum + 1.5, 9)


def write_table(
    sheet: xlsxwriter.worksheet.Worksheet,
    header: List[str],
    freeze_header: bool = False,
    format_header: xlsxwriter.format.Format | None = None,
    header_height: float | None = None,
    comment_header: List[str] | None = None,
    format_comments: dict | None = None,
    hide_columns: bool = False,
    column_widths: List[float | None] | None = None,
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
    header_height
        Header row height in character units
        (default adjusts to content, standard is 15).
    comment_header
        Whether and what text to add as a comment to each header cell (or None to skip).
    format_comments
        Whether and how to format header comments. See
        https://xlsxwriter.readthedocs.io/working_with_cell_comments.html#cell-comments.
    hide_columns
        Whether to hide unused columns.
        Ignored if `comment_header` is True, as Excel prevents hiding columns that
        overlap a comment.
    column_widths
        Width of each column in character units (or None to leave unchanged).
        Default is the minimum width to fit the header plus 1.5 character units,
        and no less than 9 character units.
        See :func:`calculate_minimum_cell_width` as a starting point for customization.
    """
    # Write header
    sheet.write_row(0, 0, header, format_header)
    # Set header height
    if header_height is not None:
        if header_height == 15:
            # Force Excel to set height (15 characters is the default)
            header_height = header_height + 1e-3
        sheet.set_row(0, height=header_height)
    # Hide unused columns
    if hide_columns and not comment_header:
        sheet.set_column(len(header), MAX_COLS - 1, options={'hidden': 1})
    # Freeze header
    if freeze_header:
        sheet.freeze_panes(1, 0)
    # Resize columns and add header comments
    format = format_header.__dict__ if format_header else {}
    for i, content in enumerate(header):
        if column_widths:
            width = column_widths[i]
        if not column_widths or width is None:
            # Determine final cell format
            width = calculate_column_width(
                header=content,
                family=format.get('font_name', 'calibri'),
                wrap=format.get('text_wrap', False),
                size=format.get('font_size', 11),
                bold=format.get('bold', False),
                italic=format.get('italic', False),
            )
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
    format_header: dict
    | None = {'bold': True, 'bg_color': '#d3d3d3', 'valign': 'top', 'text_wrap': True},
    format_comments: dict | None = {'font_size': 11, 'x_scale': 2, 'y_scale': 2},
    freeze_header: bool = True,
    header_height: float | None = None,
    hide_columns: bool = False,
    column_widths: Dict[str, List[float | None]] | None = None,
) -> xlsxwriter.Workbook | None:
    """
    Write a template for data entry to a Microsoft Excel workbook.

    Parameters
    ----------
    package
        Frictionless Data Tabular Data Package specification.
        See https://specs.frictionlessdata.io/tabular-data-package.
        Table names (`resource.name`) are used as sheet names and can be any string up
        to 31 characters long.
    path
        Path of the created Microsoft Excel (.xlsx) file.
        If None, the result is returned as :class:`xlsxwriter.Workbook`.
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
    header_height
        Header row height in character units
        (default adjusts to content, standard is 15).
    hide_columns
        Whether to hide unused columns.
    column_widths
        For each table (with `resource.name` as dictionary key),
        the widths in pixels of each column (one value per column, None to skip).
        Default is the minimum width to fit the header plus 1.5 character units,
        and no less than 9 character units.
        See :func:`calculate_minimum_cell_width` as a starting point for customization
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
            header_height=header_height,
            hide_columns=hide_columns,
            column_widths=(column_widths or {}).get(table_props['table']),
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
