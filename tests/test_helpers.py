"""Test cases for the helpers module."""
from __future__ import annotations

from typing import Any, List, Tuple

import pytest

import tablecloth.constants
import tablecloth.helpers


@pytest.mark.parametrize(
    'input, expected',
    [
        ('snakeCase', 'snake_case'),
        ('SnakeCase', 'snake_case'),
        ('snake_case', 'snake_case'),
        ('longSnakeCase', 'long_snake_case'),
        ('LongSnakeCase', 'long_snake_case'),
        ('long_snake_case', 'long_snake_case'),
    ],
)
def test_converts_camel_to_snake_case(input: str, expected: str) -> None:
    """It converts camel case string to snake case."""
    assert tablecloth.helpers.camel_to_snake_case(input) == expected


@pytest.mark.parametrize(
    'input, expected',
    [(None, []), ('', []), ('x', ['x']), (['x'], ['x']), (['x', 'y'], ['x', 'y'])],
)
def test_converts_singleton_string_to_list(
    input: str | list | None, expected: list
) -> None:
    """It converts singleton strings to a list."""
    assert tablecloth.helpers.to_list(input) == expected


@pytest.mark.parametrize(
    'input, expected',
    [
        (0, 'A'),
        (1, 'B'),
        (26, 'AA'),
        (28, 'AC'),
        (16383, 'XFD'),
    ],
)
def test_converts_column_index_to_code_and_back(input: int, expected: str) -> None:
    """It converts a zero-based column index to a spreadsheet column code (and back)."""
    code = tablecloth.helpers.column_index_to_code(input)
    assert code == expected
    assert tablecloth.helpers.column_code_to_index(code) == input


@pytest.mark.parametrize(
    'input, expected',
    [
        (0, 1),
        (1, 2),
        (99, 100),
    ],
)
def test_converts_row_index_to_code_and_back(input: int, expected: int) -> None:
    """It converts a zero-based row index to a spreadsheet row code (and back)."""
    code = tablecloth.helpers.row_index_to_code(input)
    assert code == expected
    assert tablecloth.helpers.row_code_to_index(code) == input


@pytest.mark.parametrize(
    'input, expected',
    [
        ({'col': 0, 'row': 1, 'nrows': 1}, 'A2'),
        ({'col': 0, 'row': 1, 'nrows': 1, 'fixed': True}, '$A$2'),
        ({'col': 0, 'row': 1}, 'A2:A'),
        ({'col': 0, 'row': 1, 'fixed': True}, '$A$2:$A'),
        ({'col': 0, 'row': 1, 'nrows': 2}, 'A2:A3'),
        ({'col': 0, 'row': 1, 'nrows': 2, 'fixed': True}, '$A$2:$A$3'),
        (
            {'col': 0, 'row': 1, 'nrows': 2, 'fixed': True, 'sheet': 'Sheet1'},
            "'Sheet1'!$A$2:$A$3",
        ),
        (
            {
                'col': 0,
                'row': 1,
                'nrows': 2,
                'fixed': True,
                'sheet': 'Sheet1',
                'indirect': True,
            },
            'INDIRECT("\'Sheet1\'!$A$2:$A$3")',
        ),
    ],
)
def test_builds_spreadsheet_cell_range_for_column_selection(
    input: dict, expected: str
) -> None:
    """It builds a spreadsheet cell range for a column given the selected options."""
    assert tablecloth.helpers.column_to_range(**input) == expected


@pytest.mark.parametrize(
    'input, expected',
    [
        (True, 'TRUE'),
        (False, 'FALSE'),
        (0, '0'),
        (1, '1'),
        (1.1, '1.1'),
        ('a', '"a"'),
        (None, ''),
    ],
)
def test_formats_value_for_spreadsheet_formula(
    input: bool | int | float | str | None, expected: str
) -> None:
    """It formats a value to a string for use in a spreadsheet formula."""
    assert tablecloth.helpers.format_value(input) == expected


@pytest.mark.parametrize('input', [['a']])
def test_formatting_fails_for_invalid_input(input: Any) -> None:
    """It refuses to format an invalid value."""
    with pytest.raises(ValueError):
        tablecloth.helpers.format_value(input)


@pytest.mark.parametrize(
    'input, expected',
    [
        ({'formulas': [], 'operator': 'AND'}, ''),
        ({'formulas': [], 'operator': 'OR'}, ''),
        ({'formulas': ['A2 > 0'], 'operator': 'AND'}, 'A2 > 0'),
        ({'formulas': ['A2 > 0'], 'operator': 'OR'}, 'A2 > 0'),
        ({'formulas': ['A2 > 0', 'A2 < 3'], 'operator': 'AND'}, 'AND(A2 > 0, A2 < 3)'),
        ({'formulas': ['A2 > 0', 'A2 < 3'], 'operator': 'OR'}, 'OR(A2 > 0, A2 < 3)'),
    ],
)
def test_merges_spreadsheet_formulas_by_a_logical_operator(
    input: dict, expected: str
) -> None:
    """It merges boolean spreadsheet formulas by a logical operator (AND or OR)."""
    assert tablecloth.helpers.merge_formulas(**input) == expected


@pytest.mark.parametrize(
    'input, expected',
    [
        ({'formulas': [], 'valid': True}, ''),
        ({'formulas': [], 'valid': False}, ''),
        ({'formulas': [], 'valid': True, 'ignore_blanks': []}, ''),
        ({'formulas': [], 'valid': False, 'ignore_blanks': []}, ''),
        (
            {'formulas': ['A2 > 0'], 'valid': True},
            'IF(ISBLANK({col}{row}), TRUE, A2 > 0)',
        ),
        (
            {'formulas': ['A2 > 0'], 'valid': False},
            'IF(ISBLANK({col}{row}), FALSE, A2 > 0)',
        ),
        (
            {'formulas': ['A2 > 0'], 'valid': True, 'ignore_blanks': [True]},
            'IF(ISBLANK({col}{row}), TRUE, A2 > 0)',
        ),
        (
            {'formulas': ['A2 > 0'], 'valid': False, 'ignore_blanks': [True]},
            'IF(ISBLANK({col}{row}), FALSE, A2 > 0)',
        ),
        ({'formulas': ['A2 > 0'], 'valid': True, 'ignore_blanks': [False]}, 'A2 > 0'),
        ({'formulas': ['A2 > 0'], 'valid': False, 'ignore_blanks': [False]}, 'A2 > 0'),
        # True if `A2` is null or in the interval (0, 3)
        (
            {'formulas': ['A2 > 0', 'A2 < 3'], 'valid': True},
            'IF(ISBLANK({col}{row}), TRUE, AND(A2 > 0, A2 < 3))',
        ),
        # True if `A2` is not null and in the intervals (0, ∞) or (-∞, 3)
        (
            {'formulas': ['A2 > 0', 'A2 < 3'], 'valid': False},
            'IF(ISBLANK({col}{row}), FALSE, OR(A2 > 0, A2 < 3))',
        ),
        # True if `A2` is not null and in the interval (0, 3)
        (
            {
                'formulas': ['A2 > 0', 'A2 < 3'],
                'valid': True,
                'ignore_blanks': [False, False],
            },
            'AND(A2 > 0, A2 < 3)',
        ),
        # True if `A2` is null or in the intervals (0, ∞) or (-∞, 3)
        (
            {
                'formulas': ['A2 > 0', 'A2 < 3'],
                'valid': False,
                'ignore_blanks': [False, False],
            },
            'OR(A2 > 0, A2 < 3)',
        ),
        # Contrived mixed cases
        (
            {
                'formulas': ['A2 > 0', 'A2 < 3'],
                'valid': True,
                'ignore_blanks': [False, True],
            },
            'AND(A2 > 0, IF(ISBLANK({col}{row}), TRUE, A2 < 3))',
        ),
        (
            {
                'formulas': ['A2 > 0', 'A2 < 3'],
                'valid': False,
                'ignore_blanks': [False, True],
            },
            'OR(A2 > 0, IF(ISBLANK({col}{row}), FALSE, A2 < 3))',
        ),
        (
            {
                'formulas': ['A2 <> 1', 'A2 > 0', 'A2 < 3'],
                'valid': True,
                'ignore_blanks': [False, False, True],
            },
            'AND(A2 <> 1, A2 > 0, IF(ISBLANK({col}{row}), TRUE, A2 < 3))',
        ),
        (
            {
                'formulas': ['A2 <> 1', 'A2 > 0', 'A2 < 3'],
                'valid': True,
                'ignore_blanks': [False, True, True],
            },
            'AND(A2 <> 1, IF(ISBLANK({col}{row}), TRUE, AND(A2 > 0, A2 < 3)))',
        ),
    ],
)
def test_merges_spreadsheet_conditional_formatting_formulas(
    input: dict, expected: str
) -> None:
    """It merges spreadsheet formulas used for conditional formatting."""
    assert tablecloth.helpers.merge_conditions(**input) == expected


@pytest.mark.parametrize(
    'input, expected',
    [
        ([], ''),
        (['a'], 'a'),
        (['a', 'b'], 'a and b'),
        (['a', 'b', 'c'], 'a, b, and c'),
        (['a', 'b', 'c', 'd'], 'a, b, c, and d'),
    ],
)
def test_joins_strings_into_a_human_readable_list(
    input: List[str], expected: str
) -> None:
    """It joins strings into an English grammatically correct human-readable list."""
    assert tablecloth.helpers.readable_join(input) == expected


@pytest.mark.parametrize(
    'input, expected',
    [
        ({'checks': [], 'valid': False, 'col': 'A', 'row': 2}, None),
        (
            {
                'checks': [{'formula': 'A2 > 0', 'ignore_blank': False}],
                'valid': True,
                'col': 'A',
                'row': 2,
            },
            'A2 > 0',
        ),
        (
            {
                'checks': [{'formula': 'A2 > 0', 'ignore_blank': False}],
                'valid': False,
                'col': 'A',
                'row': 2,
            },
            'A2 > 0',
        ),
        (
            {
                'checks': [{'formula': 'A2 > 0', 'ignore_blank': True}],
                'valid': True,
                'col': 'A',
                'row': 2,
            },
            'IF(ISBLANK(A2), TRUE, A2 > 0)',
        ),
        (
            {
                'checks': [{'formula': 'A2 > 0', 'ignore_blank': True}],
                'valid': False,
                'col': 'A',
                'row': 2,
            },
            'IF(ISBLANK(A2), FALSE, A2 > 0)',
        ),
        # True if `A2` is null or in the interval (0, 3)
        (
            {
                'checks': [
                    {'formula': 'A2 > 0', 'ignore_blank': True},
                    {'formula': 'A2 < 3', 'ignore_blank': True},
                ],
                'valid': True,
                'col': 'A',
                'row': 2,
            },
            'IF(ISBLANK(A2), TRUE, AND(A2 > 0, A2 < 3))',
        ),
        # True if `A2` is not null and in the intervals (0, ∞) or (-∞, 3)
        (
            {
                'checks': [
                    {'formula': 'A2 > 0', 'ignore_blank': True},
                    {'formula': 'A2 < 3', 'ignore_blank': True},
                ],
                'valid': False,
                'col': 'A',
                'row': 2,
            },
            'IF(ISBLANK(A2), FALSE, OR(A2 > 0, A2 < 3))',
        ),
        # True if `A2` is not null and in the interval (0, 3)
        (
            {
                'checks': [
                    {'formula': 'A2 > 0', 'ignore_blank': False},
                    {'formula': 'A2 < 3', 'ignore_blank': False},
                ],
                'valid': True,
                'col': 'A',
                'row': 2,
            },
            'AND(A2 > 0, A2 < 3)',
        ),
        # True if `A2` is null or in the intervals (0, ∞) or (-∞, 3)
        (
            {
                'checks': [
                    {'formula': 'A2 > 0', 'ignore_blank': False},
                    {'formula': 'A2 < 3', 'ignore_blank': False},
                ],
                'valid': False,
                'col': 'A',
                'row': 2,
            },
            'OR(A2 > 0, A2 < 3)',
        ),
    ],
)
def test_builds_conditional_formatting_formula_from_column_checks_(
    input: dict, expected: str
) -> None:
    """It builds a conditional formatting formula from column checks."""
    assert tablecloth.helpers.build_column_condition(**input) == expected


@pytest.mark.parametrize(
    'input, expected',
    [
        ([], None),
        (
            [{'formula': 'A2 > 0', 'message': '> 0'}],
            {'formula': 'A2 > 0', 'message': 'Value must be > 0', 'ignore_blank': True},
        ),
        (
            [
                {'formula': 'A2 > 0', 'message': '> 0'},
                {'formula': 'A2 < 3', 'message': '< 3'},
            ],
            {
                'formula': 'AND(A2 > 0, A2 < 3)',
                'message': 'Value must be > 0 and < 3',
                'ignore_blank': True,
            },
        ),
        (
            [
                {'formula': 'A2 > 0', 'message': '> 0'},
                {'formula': 'A2 < 3', 'message': '< 3'},
                {'formula': 'ISNUMBER(A2)', 'message': 'number'},
            ],
            {
                'formula': 'AND(A2 > 0, A2 < 3, ISNUMBER(A2))',
                'message': 'Value must be > 0, < 3, and number',
                'ignore_blank': True,
            },
        ),
    ],
)
def test_builds_data_validation_from_column_checks(
    input: List[tablecloth.constants.Check], expected: dict
) -> None:
    """It builds a data validation formula and error message from column checks."""
    assert tablecloth.helpers.build_column_validation(input) == expected


FOREIGN_KEYS = [
    {'fields': ['x'], 'reference': {'resource': 'a', 'fields': ['xx']}},
    {'fields': ['y'], 'reference': {'resource': 'a', 'fields': ['yy']}},
    {'fields': ['y', 'x'], 'reference': {'resource': 'b', 'fields': ['y', 'x']}},
]


@pytest.mark.parametrize(
    'input, expected',
    [
        ({'foreign_keys': [], 'table': 'a', 'column': 'x'}, []),
        (
            {'foreign_keys': FOREIGN_KEYS, 'table': 'a', 'column': 'x'},
            [(None, 'xx'), ('b', 'x')],
        ),
        (
            {'foreign_keys': FOREIGN_KEYS, 'table': 'a', 'column': 'y'},
            [(None, 'yy'), ('b', 'y')],
        ),
        (
            {'foreign_keys': FOREIGN_KEYS, 'table': 'b', 'column': 'x'},
            [('a', 'xx'), (None, 'x')],
        ),
        (
            {'foreign_keys': FOREIGN_KEYS, 'table': 'b', 'column': 'y'},
            [('a', 'yy'), (None, 'y')],
        ),
        ({'foreign_keys': FOREIGN_KEYS, 'table': 'b', 'column': 'z'}, []),
        (
            {'foreign_keys': FOREIGN_KEYS, 'table': 'c', 'column': 'x'},
            [('a', 'xx'), ('b', 'x')],
        ),
        (
            {'foreign_keys': FOREIGN_KEYS, 'table': 'c', 'column': 'y'},
            [('a', 'yy'), ('b', 'y')],
        ),
    ],
)
def test_reduces_foreign_keys_to_all_simple_keys_for_one_column(
    input: dict, expected: List[Tuple[str | None, str]]
) -> None:
    """It reduces foreign keys to the unique set of simple foreign keys for a column."""
    assert tablecloth.helpers.reduce_foreign_keys(**input) == expected
