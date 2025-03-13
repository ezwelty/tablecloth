"""Test cases for the layout module."""
import pytest

from tablecloth import Layout
from tablecloth.constants import Constraints


def test_gets_a_column_code() -> None:
    """It gets a column code."""
    layout = Layout()
    layout.set_table('a', ['x', 'y'])
    assert layout.get_column_code('a', 'x') == 'A'
    assert layout.get_column_code('a', 'y') == 'B'


def test_gets_a_column_range() -> None:
    """It gets a column range."""
    layout = Layout()
    layout.set_table('a', ['x', 'y'])
    # First column
    assert layout.get_column_range('a', 'x') == 'A2:A'
    assert layout.get_column_range('a', 'x', nrows=1) == 'A2'
    assert layout.get_column_range('a', 'x', nrows=1, fixed=True) == '$A$2'
    assert layout.get_column_range('a', 'x', fixed=True) == '$A$2:$A'
    assert layout.get_column_range('a', 'x', absolute=True) == "'a'!A2:A"
    assert (
        layout.get_column_range('a', 'x', absolute=True, indirect=True)
        == 'INDIRECT("\'a\'!A2:A")'
    )
    # Second column
    assert layout.get_column_range('a', 'y') == 'B2:B'
    # Sheet name different from table name
    layout.tables[0]['sheet'] = 'Sheet1'
    assert layout.get_column_range('a', 'x', absolute=True) == "'Sheet1'!A2:A"
    # With max_rows
    layout.max_rows = 10
    assert layout.get_column_range('a', 'x') == 'A2:A10'
    assert layout.get_column_range('a', 'x', fixed=True) == '$A$2:$A$10'


def test_gets_an_enum_range() -> None:
    """It gets an enum range."""
    layout = Layout()
    # First enum
    enum = [1, 2]
    layout.set_enum(enum)
    assert layout.get_enum_range(enum) == "'lists'!$A$1:$A$2"
    assert (
        layout.get_enum_range(enum, indirect=True) == 'INDIRECT("\'lists\'!$A$1:$A$2")'
    )
    # Second enum
    enum = [1, 2, 3]
    layout.set_enum(enum)
    assert layout.get_enum_range(enum) == "'lists'!$B$1:$B$3"
    assert (
        layout.get_enum_range(enum, indirect=True) == 'INDIRECT("\'lists\'!$B$1:$B$3")'
    )
    # Non-default enum sheet name
    layout.enum_sheet = 'Sheet1'
    assert layout.get_enum_range(enum) == "'Sheet1'!$B$1:$B$3"


def test_fails_if_table_already_exists() -> None:
    """It fails if a table with the same name already exists."""
    layout = Layout()
    layout.set_table(table='a', columns=['x'], sheet='a')
    with pytest.raises(ValueError):
        layout.set_table(table='a', columns=['y'], sheet='a')
    with pytest.raises(ValueError):
        layout.set_table(table='a', columns=['y'], sheet='b')


def test_fails_if_sheet_already_exists() -> None:
    """It fails if a sheet with the same name already exists."""
    layout = Layout()
    layout.set_table('a', ['x'], sheet='sheet')
    with pytest.raises(ValueError):
        layout.set_table('b', ['y'], sheet='sheet')


def test_fails_if_sheet_name_too_long() -> None:
    """It fails if a sheet name is too long."""
    layout = Layout(max_name_length=2)
    with pytest.raises(ValueError):
        layout.set_table('a', ['x'], sheet='xyz')


def test_fails_if_column_names_are_not_unique() -> None:
    """It fails if column names are not unique."""
    layout = Layout()
    with pytest.raises(ValueError):
        layout.set_table('a', ['x', 'x'])


def test_fails_if_table_not_found() -> None:
    """It fails if no table exists with that name."""
    layout = Layout()
    with pytest.raises(KeyError):
        layout.get_table('a')
    layout.set_table('a', ['x'])
    with pytest.raises(KeyError):
        layout.get_table('b')


def test_fails_if_enum_not_found() -> None:
    """It fails if no enum exists with those values."""
    layout = Layout()
    with pytest.raises(KeyError):
        layout.get_enum([1, 2])
    layout.set_enum([1, 2])
    with pytest.raises(KeyError):
        layout.get_enum([2, 1])


def test_fails_if_enum_empty() -> None:
    """It fails if an enum is empty."""
    layout = Layout()
    with pytest.raises(ValueError):
        layout.set_enum([])


def test_warns_if_enum_value_starts_with_special_characters() -> None:
    """It warns if an enum value starts with special spreadsheet characters."""
    layout = Layout()
    with pytest.warns(UserWarning):
        layout.set_enum(['+'])
    with pytest.warns(UserWarning):
        layout.set_enum(['='])
    with pytest.warns(UserWarning):
        layout.set_enum(["'"])


def test_sets_same_enum_multiple_times() -> None:
    """It sets the same enum multiple times without duplication."""
    layout = Layout()
    layout.set_enum([1, 2])
    layout.set_enum([1, 2])
    layout.set_enum([1, 2])
    assert len(layout.enums) == 1


@pytest.mark.parametrize('constraint', ['minimum', 'maximum', 'minLength', 'maxLength'])
def test_adds_check_for_constraint_equal_to_zero(constraint: str) -> None:
    """It adds a check for a constraint equal to zero."""
    table, column = 'a', 'x'
    layout = Layout()
    layout.set_table(table, [column])
    checks = layout.gather_column_checks(
        table,
        column,
        valid=True,
        # HACK: Cast to Constraints to avoid mypy error
        constraints=Constraints(**{constraint: 0}),
    )
    assert checks
