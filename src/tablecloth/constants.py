from __future__ import annotations

from typing import Dict, List, Literal, TypedDict


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


class Dropdown(TypedDict):
    """Column dropdown."""

    values: List[str] | str
    """Dropdown values as a list or cell range."""
    source: Literal['boolean', 'foreign_key', 'enum']
    """Source of dropdown values."""


class ForeignKeyReference(TypedDict):
    """Foreign key reference."""

    resource: str | None
    """Table name."""
    fields: str | List[str]
    """Column name(s)."""


class ForeignKey(TypedDict):
    """Foreign key (https://specs.frictionlessdata.io/table-schema/#foreign-keys)."""

    fields: str | List[str]
    """Local column name(s)."""
    reference: ForeignKeyReference
    """Foreign reference."""


class Constraints(TypedDict, total=False):
    """
    Column constraints (https://specs.frictionlessdata.io/table-schema/#constraints).
    """

    required: bool
    """Whether a value is required."""
    unique: bool
    """Whether values must be unique."""
    min_length: int
    """Minimum length (of string)."""
    max_length: int
    """Maximum length (of string)."""
    minimum: int | float
    """Minimum value."""
    maximum: int | float
    """Maximum value."""
    pattern: str
    """Regular expression."""
    enum: List[int] | List[float] | List[str]
    """List of allowed values."""


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

Supports a subset of the types in the Frictionless Data Table Schema specification:
https://specs.frictionlessdata.io/table-schema/#types-and-formats

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

Supports all constraints in the Frictionless Data Table Schema specification:
https://specs.frictionlessdata.io/table-schema/#constraints
(`enum` is handled separately).

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
