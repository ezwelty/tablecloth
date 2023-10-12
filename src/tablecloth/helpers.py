"""Helper functions."""
from __future__ import annotations

import functools
import re
from typing import List, Literal, Tuple

from . import constants

# ---- Functions ----


def camel_to_snake_case(x: str) -> str:
    """Convert camelCase (and CamelCase) to snake_case."""
    pattern = re.compile(r'(?<!^)(?=[A-Z])')
    return pattern.sub('_', x).lower()


def to_list(x: str | list | None) -> list:
    """Cast to list, wrapping singleton string as first element."""
    if not x:
        return []
    if isinstance(x, str):
        return [x]
    return list(x)


def column_index_to_code(i: int) -> str:
    """Convert a column index to a spreadsheet column code."""
    letters: List[str] = []
    i = i + 1
    while i:
        i, remainder = divmod(i - 1, 26)
        letters[:0] = constants.LETTERS[remainder]
    return ''.join(letters)


def column_code_to_index(code: str) -> int:
    """Convert a spreadsheet column code to a column index (zero-based)."""

    # https://gist.github.com/dbspringer/643254008e6784aa749e#file-col2num-py
    def function(x: int, y: int) -> int:
        return x * 26 + y

    return functools.reduce(function, [ord(c) - ord('A') + 1 for c in code]) - 1


def row_index_to_code(i: int) -> int:
    """Convert a row index (zero-based) to a spreadsheet row code."""
    return i + 1


def row_code_to_index(code: int) -> int:
    """Convert a spreadsheet row code to a row index (zero-based)."""
    return code - 1


def column_to_range(
    col: int,
    row: int,
    nrows: int | None = None,
    fixed: bool = False,
    sheet: str | None = None,
    indirect: bool = False,
) -> str:
    """
    Get a column's cell range in spreadsheet notation.

    Parameters
    ----------
    col
        Column index (zero-based).
    row
        Start row index (zero-based).
    nrows
        Number of rows to include (or unbounded if None).
    fixed
        Whether to use a fixed range (e.g. $A$2:$A).
    sheet
        Whether to refer to the range by sheet name (e.g. 'Sheet1'!A2:A).
    indirect
        Whether to wrap the range in the `INDIRECT` function.
        See https://support.google.com/docs/answer/3093377.
        Ignored if `sheet` is None.
    """
    col_code = column_index_to_code(col)
    row_code = row_index_to_code(row)
    prefix = '$' if fixed else ''
    cells = f'{prefix}{col_code}{prefix}{row_code}'
    if nrows == 1:
        return cells
    cells += f':{prefix}{col_code}'
    if nrows is not None:
        last_row = row + nrows
        cells += f'{prefix}{last_row}'
    if sheet:
        cells = f"'{sheet}'!{cells}"
        if indirect:
            cells = f'INDIRECT("{cells}")'
    return cells


def format_value(x: bool | int | float | str | None) -> str:
    """
    Format a singleton value to a string for spreadsheet cells or formulas.

    Parameters
    ----------
    x
        Singleton value.

    Raises
    ------
    ValueError
        Unexpected value `x` with unsupported type.
    """
    if isinstance(x, bool):
        return str(x).upper()
    if isinstance(x, (int, float)):
        return str(x)
    if isinstance(x, str):
        return f'"{x}"'
    if x is None:
        return ''
    raise ValueError(f'Unexpected value {x} with unsupported type {type(x)}')


def merge_formulas(formulas: List[str], operator: Literal['AND', 'OR']) -> str:
    """
    Merge formulas (returning TRUE or FALSE) by a logical operator.

    Parameters
    ----------
    formulas
        Spreadsheet formulas.
    operator
        Logical operator.
    """
    if not formulas:
        return ''
    if len(formulas) == 1:
        return formulas[0]
    return f"{operator}({', '.join(formulas)})"


def merge_conditions(
    formulas: List[str], valid: bool, ignore_blanks: List[bool] | None = None
) -> str:
    """
    Merge formulas (returning TRUE or FALSE) for use in conditional formatting.

    Parameters
    ----------
    formulas
        Formulas returning a boolean value (TRUE or FALSE).
    valid
        Whether all `formulas` return TRUE for valid (`True`) or
        invalid (`False`) values.
    ignore_blanks
        Whether each formula ignores blank values (assumed by default).

    Returns
    -------
    formula :
        Merged formula.
        All `formulas` which `ignore_blanks` are wrapped in an if statement
        with column and row placeholders (`IF(ISBLANK({col}{row}), ...`).
    """
    operator: Literal['AND', 'OR'] = 'AND' if valid else 'OR'
    fs, fs_ignore = [], []
    for i, formula in enumerate(formulas):
        if ignore_blanks and not ignore_blanks[i]:
            fs.append(formula)
        else:
            fs_ignore.append(formula)
    if fs_ignore:
        # Wrap formulas ignoring blank in single if statement
        merged = merge_formulas(fs_ignore, operator=operator)
        fs.append(f'IF(ISBLANK({{col}}{{row}}), {format_value(valid)}, {merged})')
    return merge_formulas(fs, operator=operator)


def build_column_condition(
    checks: List[constants.Check], valid: bool, col: str, row: int = 2
) -> str | None:
    """
    Build a column's conditional formatting formula from column checks.

    Assumes that blank cells are not ignored by the application,
    which is the case for Microsoft Excel and Google Sheets.

    Parameters
    ----------
    checks
        Column checks.
    valid
        Whether `formulas` return TRUE if the value is valid (True) or invalid (False).
    col
        Column code.
    row
        Start row code (1-based).

    Returns
    -------
    formula :
        Formula that returns TRUE if all formulas return TRUE (`valid` is True),
        or TRUE if any formula returns TRUE (`valid` is False).
    """
    if not checks:
        return None
    formula = merge_conditions(
        formulas=[x['formula'] for x in checks],
        valid=valid,
        ignore_blanks=[x['ignore_blank'] for x in checks],
    )
    return formula.format(col=col, row=row)


def readable_join(values: List[str]) -> str:
    """Return an English grammatically-correct list (with an Oxford comma)."""
    if len(values) < 3:
        return ' and '.join(values)
    return ', '.join(values[:-1]) + ', and ' + values[-1]


def build_column_validation(checks: List[constants.Check]) -> constants.Check | None:
    """
    Build a column's validation from column checks.

    Assumes that blank cells are ignored by the application,
    which is the case for Google Sheets and the default for Microsoft Excel.

    Parameters
    ----------
    checks
        Column checks. Formulas should return TRUE if value is valid.
    """
    if not checks:
        return None
    formulas = [x['formula'] for x in checks]
    messages = [x['message'] for x in checks]
    return {
        'formula': merge_formulas(formulas, operator='AND'),
        'message': f'Value must be {readable_join(messages)}',
        'ignore_blank': True,
    }


def reduce_foreign_keys(
    foreign_keys: List[constants.ForeignKey], table: str, column: str
) -> List[Tuple[str | None, str]]:
    """Reduce foreign keys to the unique set of simple foreign keys for one column."""
    keys = []
    for foreign_key in foreign_keys or []:
        columns = to_list(foreign_key['fields'])
        foreign_columns = to_list(foreign_key['reference']['fields'])
        try:
            i = columns.index(column)
        except ValueError:
            continue
        foreign_column = foreign_columns[i]
        foreign_table = foreign_key['reference'].get('resource') or table
        key = (None if foreign_table == table else foreign_table, foreign_column)
        if key not in keys:
            keys.append(key)
    return keys
