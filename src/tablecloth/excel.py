from collections import defaultdict
from pathlib import Path
from typing import Dict, List, Literal, Optional, Union

try:
    import xlsxwriter
    import xlsxwriter.format
    import xlsxwriter.worksheet
except ImportError:
    raise ImportError('Writing Excel templates requires `xlsxwriter`')

from . import constants
from . import helpers
from .layout import Layout


MAX_COLS: int = 16384
"""Maximum number of columns per sheet."""

MAX_COLS: int = 16384
"""Maximum number of columns per sheet."""

MAX_NAME_LENGTH: int = 31
"""Maximum length of sheet name."""

MAX_ROWS: int = 1048576
"""Maximum length of sheet name."""


def write_table(
    sheet: xlsxwriter.worksheet.Worksheet,
    header: List[str],
    freeze_header: bool = False,
    format_header: xlsxwriter.format.Format = None,
    comment_header: List[str] = None,
    format_comments: dict = None,
    hide_columns: bool = False,
) -> None:
    """
    Write empty table to Microsoft Excel.

    Parameters
    ----------
    sheet
        Spreadsheet.
    header
        Content of the first (header) row.
    freeze_header
        Whether to freeze the header.
    format_header
        Whether and how to format the header.
    comment_header
        Whether and what text to add to the header.
    format_comments
        Whether and how to format header comments.
    hide_columns
        Whether to hide unused columns.
        Ignored if `comment_header`, as Excel prevents hiding columns that
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
    Write enumerated values to Microsoft Excel.

    Parameters
    ----------
    sheet
        Spreadsheet.
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
    path: Union[str, Path] = None,
    enum_sheet: str = 'lists',
    hide_enum_sheet: bool = True,
    header_comments: Dict[str, List[str]] = None,
    dropdowns: bool = True,
    error_type: Literal['information', 'warning', 'stop'] = None,
    format_invalid: Optional[dict] = {'bg_color': '#ffc7ce'},
    format_header: Optional[dict] = {'bold': True, 'bg_color': '#d3d3d3'},
    format_comments: Optional[dict] = {'font_size': 11, 'x_scale': 2, 'y_scale': 2},
    freeze_header: bool = True,
    hide_columns: bool = False,
) -> Optional[xlsxwriter.Workbook]:
    """
    Write package template.

    Parameters
    ----------
    error_type
      * stop: Display error message with buttons to cancel or retry
      * warning: Display error message with buttons to accept or retry
      * information: Display error message with button to accept
    """
    # ---- Initialize
    layout = Layout(
        enum_sheet=enum_sheet,
        max_name_length=MAX_NAME_LENGTH,
        max_rows=MAX_ROWS,
        indirect=False,
    )
    book = xlsxwriter.Workbook(filename=path)
    # Register formats
    if format_invalid:
        format_invalid: xlsxwriter.format.Format = book.add_format(format_invalid)
    if format_header:
        format_header: xlsxwriter.format.Format = book.add_format(format_header)

    # ---- Write tables
    for resource in package['resources']:
        table = resource['name']
        columns = [field['name'] for field in resource['schema']['fields']]
        layout.set_table(table, columns)
        sheet = book.add_worksheet(table)
        write_table(
            sheet,
            header=[field['name'] for field in resource['schema']['fields']],
            comment_header=header_comments.get(table) if header_comments else None,
            format_header=format_header,
            format_comments=format_comments,
            freeze_header=freeze_header,
            hide_columns=hide_columns,
        )

    # --- Add column checks
    for resource in package['resources']:
        table = resource['name']
        sheet: xlsxwriter.worksheet.Worksheet = book.get_worksheet_by_name(table)

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
            cells = layout.get_column_range(table, column)
            dtype = field.get('type', 'any')
            enum = field.get('constraints', {}).get('enum')
            constraints = {
                helpers.camel_to_snake_case(k): v
                for k, v in field.get('constraints', {}).items()
                if (
                    helpers.camel_to_snake_case(k) in constants.CONSTRAINTS
                    and
                    # No regex support in Excel
                    k not in ['pattern']
                )
            }

            # Register enum
            if enum and (dropdowns or error_type or format_invalid):
                layout.set_enum(enum)
                esheet = book.get_worksheet_by_name(enum_sheet)
                if not esheet:
                    esheet = book.add_worksheet(enum_sheet)
                    esheet.protect()
                write_enum(sheet=esheet, values=enum, col=layout.get_enum(enum)['col'])

            # Data validation
            validation = None
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
                        indirect=layout.indirect,
                    )
                elif enum:
                    values = layout.get_enum_range(enum)
                if values:
                    validation = {
                        'validate': 'list',
                        'value': values,
                        'error_title': 'Invalid value',
                        'error_message': 'Value must be in dropdown list',
                        'ignore_blank': True,
                        'error_type': error_type or 'information',
                        'show_error': bool(error_type),
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
                    foreign_keys=foreign_keys.get(column)
                )
                validation = helpers.build_column_validation(checks)
                if validation:
                    validation = {
                        'validate': 'custom',
                        'value': validation['formula'],
                        'error_title': 'Invalid value',
                        'error_message': validation['message'],
                        'ignore_blank': validation['ignore_blank'],
                        'error_type': error_type,
                        'show_error': bool(error_type),
                    }
            if validation:
                cells = layout.get_column_range(table, column)
                sheet.data_validation(cells, validation)

            # Conditional formatting
            if format_invalid:
                checks = layout.gather_column_checks(
                    table,
                    column,
                    valid=False,
                    dtype=dtype,
                    **constraints,
                    enum=enum,
                    foreign_keys=foreign_keys.get(column)
                )
                formula = helpers.build_column_condition(checks, valid=False)
                if formula:
                    formula = formula.format(
                        col=layout.get_column_code(table, column), row=2
                    )
                    sheet.conditional_format(
                        cells,
                        {
                            'type': 'formula',
                            # NOTE: Move string formatting elsewhere?
                            'criteria': formula,
                            'format': format_invalid,
                        },
                    )
    if hide_enum_sheet:
        sheet = book.get_worksheet_by_name(enum_sheet)
        if sheet:
            sheet.hide()
    if path is None:
        return book
    book.close()
