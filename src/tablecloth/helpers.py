import functools
from typing import List, Literal, Optional, Union
import re

from . import constants


# ---- Functions ----


def camel_to_snake_case(x: str) -> str:
    """
    Convert camelCase (and CamelCase) to snake_case.

    Examples:
    >>> camel_to_snake_case('camelCase')
    'camel_case'
    >>> camel_to_snake_case('CamelCase')
    'camel_case'
    >>> camel_to_snake_case('snake_case')
    'snake_case'
    """
    pattern = re.compile(r'(?<!^)(?=[A-Z])')
    return pattern.sub('_', x).lower()


def to_list(x: Optional[Union[str, list]]) -> list:
    """
    Cast to list, wrapping singleton string as first element.

    Examples
    --------
    >>> to_list(None)
    []
    >>> to_list('')
    []
    >>> to_list('x')
    ['x']
    >>> to_list(['x'])
    ['x']
    """
    if not x:
        return []
    if isinstance(x, str):
        return [x]
    return list(x)


def column_index_to_code(i: int) -> str:
    """
    Convert a column index to a spreadsheet column code.

    Examples
    --------
    >>> column_index_to_code(0)
    'A'
    >>> column_index_to_code(26)
    'AA'
    >>> column_index_to_code(16383)
    'XFD'
    >>> i = 1
    >>> i == column_code_to_index(column_index_to_code(i))
    True
    """
    letters: List[str] = []
    i = i + 1
    while i:
        i, remainder = divmod(i - 1, 26)
        letters[:0] = constants.LETTERS[remainder]
    return ''.join(letters)


def column_code_to_index(code: str) -> int:
    """
    Convert a spreadsheet column code to a column index.

    Examples
    --------
    >>> column_code_to_index('A')
    0
    >>> column_code_to_index('AA')
    26
    >>> column_code_to_index('XFD')
    16383
    >>> code = 'B'
    >>> code == column_index_to_code(column_code_to_index(code))
    True
    """
    # https://gist.github.com/dbspringer/643254008e6784aa749e#file-col2num-py
    function = lambda x, y: x * 26 + y
    return functools.reduce(function, [ord(c) - ord('A') + 1 for c in code]) - 1


def row_index_to_code(i: int) -> int:
    """
    Convert a row index to a spreadsheet row code.

    Examples
    --------
    >>> i = 1
    >>> row_index_to_code(i)
    2
    >>> i == row_code_to_index(row_index_to_code(i))
    True
    """
    return i + 1


def row_code_to_index(code: int) -> int:
    """
    Convert a spreadsheet row code to a row index.

    Examples
    --------
    >>> i = 2
    >>> row_code_to_index(i)
    1
    >>> i == row_index_to_code(row_code_to_index(i))
    True
    """
    return code - 1


def column_to_range(
    col: int, row: int, nrows: int = None, fixed: bool = False, sheet: str = None
) -> str:
    """
    Examples
    --------
    >>> column_to_range(0, 1, nrows=1)
    'A2'
    >>> column_to_range(0, 1, nrows=1, fixed=True)
    '$A$2'
    >>> column_to_range(0, 1)
    'A2:A'
    >>> column_to_range(0, 1, fixed=True)
    '$A$2:$A'
    >>> column_to_range(0, 1, nrows=2)
    'A2:A3'
    >>> column_to_range(0, 1, nrows=2, fixed=True)
    '$A$2:$A$3'
    >>> column_to_range(0, 1, nrows=2, fixed=True, sheet='Sheet1')
    "'Sheet1'!$A$2:$A$3"
    """
    col: str = column_index_to_code(col)
    row: str = row_index_to_code(row)
    prefix = '$' if fixed else ''
    cells = f'{prefix}{col}{prefix}{row}'
    if nrows == 1:
        return cells
    cells += f':{prefix}{col}'
    if nrows is not None:
        last_row = row + nrows - 1
        cells += f'{prefix}{last_row}'
    if sheet:
        cells = f"'{sheet}'!{cells}"
    return cells


def format_value(x: Union[bool, int, float, str, None]) -> str:
    """
    Format a singleton value for spreadsheet cells or formulas.

    Examples
    --------
    >>> format_value(True)
    'TRUE'
    >>> format_value(1)
    '1'
    >>> format_value('a')
    '"a"'
    >>> format_value(None)
    ''
    """
    if isinstance(x, bool):
        return str(x).upper()
    if isinstance(x, (int, float)):
        return str(x)
    if isinstance(x, str):
        return f'"{x}"'
    if x is None:
        return ''
    raise ValueError(f'Unexpected value {x} of type {type(x)}')


def merge_formulas(formulas: List[str], operator: Literal['AND', 'OR']) -> str:
    """
    Merge formulas (returning TRUE or FALSE) by a logical operator.

    Examples
    --------
    >>> merge_formulas([], 'AND')
    ''
    >>> merge_formulas(['A2 > 0'], 'AND')
    'A2 > 0'
    >>> merge_formulas(['A2 > 0', 'A2 < 3'], 'AND')
    'AND(A2 > 0, A2 < 3)'
    >>> merge_formulas(['A2 > 0', 'A2 < 3'], 'OR')
    'OR(A2 > 0, A2 < 3)'
    """
    if not formulas:
        return ''
    if len(formulas) == 1:
        return formulas[0]
    return f"{operator}({', '.join(formulas)})"


def merge_conditions(
    formulas: List[str], valid: bool, ignore_blanks: List[bool] = None
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
    str
        Merged formula.
        All `formulas` which `ignore_blanks` are wrapped in an if statement
        with column and row placeholders (`IF(ISBLANK({col}{row}), ...`).

    Examples
    --------
    >>> formulas = ['A2 > 0', 'A2 < 3']

    True if `A2` is null or in the interval (0, 3):

    >>> merge_conditions(formulas, valid=True)
    'IF(ISBLANK({col}{row}), TRUE, AND(A2 > 0, A2 < 3))'

    True if `A2` is not null and in the interval (0, 3):

    >>> merge_conditions(formulas, valid=True, ignore_blanks=[False, False])
    'AND(A2 > 0, A2 < 3)'

    True if `A2` is not null and in the intervals (0, ∞) or (-∞, 3):

    >>> merge_conditions(formulas, valid=False)
    'IF(ISBLANK({col}{row}), FALSE, OR(A2 > 0, A2 < 3))'

    True if `A2` is null or in the intervals (0, ∞) or (-∞, 3):

    >>> merge_conditions(formulas, valid=False, ignore_blanks=[False, False])
    'OR(A2 > 0, A2 < 3)'

    >>> merge_conditions(formulas, valid=True, ignore_blanks=[False, True])
    'AND(A2 > 0, IF(ISBLANK({col}{row}), TRUE, A2 < 3))'
    >>> merge_conditions(formulas, valid=False, ignore_blanks=[False, True])
    'OR(A2 > 0, IF(ISBLANK({col}{row}), FALSE, A2 < 3))'
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


def build_column_condition(checks: List[dict], valid: bool) -> Optional[str]:
    """
    Build a column's conditional formatting formula from column checks.

    Assumes that blank cells are not ignored by the application,
    which is the case for Microsoft Excel and Google Sheets.

    Parameters
    ----------
    checks
      Column checks.
    valid
      Whether formulas return TRUE if the value is valid or invalid.
      Returned formula will return TRUE if all formulas return TRUE (valid),
      or TRUE if any formula returns TRUE (invalid).
    """
    if not checks:
        return None
    formulas = [x['formula'] for x in checks]
    ignore_blanks = [x['ignore_blank'] for x in checks]
    return merge_conditions(formulas, valid=valid, ignore_blanks=ignore_blanks)


def build_column_validation(checks: List[dict]) -> Optional[constants.Check]:
    """
    Build a column's validation from column checks.

    Assumes that blank cells are ignored by the application,
    which is the case for Google Sheets and the default for Microsoft Excel.

    Parameters
    ----------
    checks:
      Column checks. Formulas should return TRUE if value is valid.
    """
    if not checks:
        return None
    formulas = [x['formula'] for x in checks]
    messages = [x['message'] for x in checks]
    return {
        'formula': merge_formulas(formulas, operator='AND'),
        'message': constants.JOIN_CHECK_MESSAGES(messages),
        'ignore_blank': True,
    }
