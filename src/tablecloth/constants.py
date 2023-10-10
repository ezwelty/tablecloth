from typing import Dict, List, TypedDict


class Table(TypedDict):
    """Table properties."""

    sheet: str
    """Sheet name."""
    table: str
    """Table name."""
    columns: List[str]
    """Column names."""


class Enum(TypedDict):
    """Enum properties."""

    values: list
    """Enum values."""
    col: int
    """Column (zero-indexed)."""


class Check(TypedDict):
    """Formula for checking column values."""

    formula: str
    """Formula returning TRUE or FALSE."""
    message: str
    """Message describing the check."""
    ignore_blank: bool
    """Whether blank cells can be ignored."""


class CheckTemplate(TypedDict):
    """Formula templates for checking column values."""

    valid: str
    """Formula returning TRUE if valid."""
    invalid: str
    """Formula returning TRUE if invalid."""
    message: str
    """Message describing the check."""
    ignore_blank: bool
    """Whether to skip the checking of blank cells."""


LETTERS: str = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
"""Letters of the latin alphabet."""


TYPES: Dict[str, CheckTemplate] = {
    'number': {
        'valid': 'ISNUMBER({col}{row})',
        'invalid': 'NOT(ISNUMBER({col}{row}))',
        'message': 'number',
        'ignore_blank': True,
    },
    'integer': {
        'valid': 'IF(ISNUMBER({col}{row}), INT({col}{row}) = {col}{row}, FALSE)',
        'invalid': 'IF(ISNUMBER({col}{row}), INT({col}{row}) <> {col}{row}, TRUE)',
        'message': 'integer',
        'ignore_blank': True,
    },
    'year': {
        'valid': 'IF(ISNUMBER({col}{row}), INT({col}{row}) = {col}{row}, FALSE)',
        'invalid': 'IF(ISNUMBER({col}{row}), INT({col}{row}) <> {col}{row}, TRUE)',
        'message': 'year',
        'ignore_blank': True,
    },
    'boolean': {
        'valid': 'OR({col}{row} = TRUE, {col}{row} = FALSE)',
        'invalid': 'AND({col}{row} <> TRUE, {col}{row} <> FALSE)',
        'message': 'TRUE or FALSE',
        'ignore_blank': True,
    },
}
"""
Formula templates for (in)valid type checks.

* col: Column code.
* row: First (minimum) row number.
"""


CONSTRAINTS: Dict[str, CheckTemplate] = {
    'required': {
        'valid': 'NOT(ISBLANK({col}{row}))',
        'invalid': (
            'AND(ISBLANK({col}{row}), '
            'COUNTBLANK(${min_col}{row}:${max_col}{row}) <> {ncols})'
        ),
        'message': 'not blank',
        'ignore_blank': False,
    },
    'unique': {
        'valid': 'COUNTIF({col}${row}:{col}{max_row}, {col}{row}) < 2',
        'invalid': 'COUNTIF({col}${row}:{col}{max_row}, {col}{row}) >= 2',
        'message': 'unique',
        'ignore_blank': True,
    },
    'min_length': {
        'valid': 'LEN({col}{row}) >= {value}',
        'invalid': 'LEN({col}{row}) < {value}',
        'message': 'length ≥ {value}',
        'ignore_blank': True,
    },
    'max_length': {
        'valid': 'LEN({col}{row}) <= {value}',
        'invalid': 'LEN({col}{row}) > {value}',
        'message': 'length ≤ {value}',
        'ignore_blank': True,
    },
    'minimum': {
        'valid': '{col}{row} >= {value}',
        'invalid': '{col}{row} < {value}',
        'message': '≥ {value}',
        'ignore_blank': True,
    },
    'maximum': {
        'valid': '{col}{row} <= {value}',
        'invalid': '{col}{row} > {value}',
        'message': '≤ {value}',
        'ignore_blank': True,
    },
    'pattern': {
        'valid': 'REGEXMATCH(TO_TEXT({col}{row}), "^{value}$")',
        'invalid': 'NOT(REGEXMATCH(TO_TEXT({col}{row}), "^{value}$"))',
        'message': 'matching the regular expression {value}',
        'ignore_blank': True,
    },
}
"""
Formula templates for (in)valid constraint checks.

* col: Column code.
* row: First (minimum) row number.
* value: Constraint value.
* min_col: Code of first column in table.
* max_col: Code of last column in table.
* max_row: Last row number (optional), prefixed as needed by '$'.
* ncols: Number of columns in table.
"""


IN_RANGE: CheckTemplate = {
    'valid': 'ISNUMBER(MATCH({col}{row}, {range}, 0))',
    'invalid': 'ISNA(MATCH({col}{row}, {range}, 0))',
    'message': 'in the range {range}',
    'ignore_blank': True,
}
"""
Formula templates to check whether a value is in a range.

* col: Column code.
* row: First (minimum) row number.
* range: Cell range to look in.
"""
