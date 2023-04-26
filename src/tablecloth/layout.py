import re
from typing import List, Literal, Tuple, Union
import warnings

from . import constants
from . import helpers


class Layout:
  """
  Tabular spreadsheet layout.

  Attributes
  ----------
  tables
      Tables added by :meth:`set_table`.
  enums
    Enums added by :meth:`set_enum`.
  max_rows
    Maximum number of rows allowed per sheet.
    If `None`, column ranges are unbounded (e.g. `A2:A`).
    Otherwise, they are bounded (e.g. `A2:A1000`).
  max_name_length:
    Maximum length of sheet name (in characters).
  indirect:
    Whether to wrap foreign table lookups in `INDIRECT` function.
    See https://support.google.com/docs/answer/3093377.
  """

  def __init__(
    self,
    enum_sheet: str = 'lists',
    max_rows: int = None,
    max_name_length: int = None,
    indirect: bool = False
  ) -> None:
    self.tables: List[constants.Table] = []
    self.enums: List[constants.Enum] = []
    self.enum_sheet = enum_sheet
    self.max_rows = max_rows
    self.max_name_length = max_name_length
    self.indirect = indirect

  def set_table(
    self,
    table: str,
    columns: List[str],
    sheet: str = None
  ) -> None:
    """
    Add a new table to a new sheet.

    Parameters
    ----------
    table
        Table name. Must be unique.
    columns
        Column names. Cannot contain duplicates.
    sheet
        Sheet name (`table` by default).
        Must not already be taken by a table or enum.
    """
    sheet = sheet or table
    if table in [x['table'] for x in self.tables]:
        raise ValueError(f"Table '{table}' already exists")
    if sheet in [x['sheet'] for x in self.tables] or sheet == self.enum_sheet:
        raise ValueError(f"Sheet '{sheet}' already exists")
    if self.max_name_length and len(sheet) > self.max_name_length:
        raise ValueError(
          f"Sheet name cannot be longer than {self.max_name_length}"
        )
    if len(columns) != len(set(columns)):
        raise ValueError(f'Columns {columns} are not unique')
    self.tables.append({'sheet': sheet, 'table': table, 'columns': columns})

  def get_table(self, table: str) -> dict:
    """Get table properties by name."""
    for x in self.tables:
      if x['table'] == table:
        return x
    raise KeyError(f"Table '{table}' not found")

  def set_enum(self, values: list) -> None:
    """Add an enum (if new) to the enum sheet."""
    formulas = [
      x for x in values
      if isinstance(x, str) and re.match(r"^[+=']", x)
    ]
    if formulas:
      warnings.warn(
        f"Enum values start with +, =, or ' {formulas}. "
        'Expect unexpected behavior.'
      )
    if values in [x['values'] for x in self.enums]:
      return
    col = max([0, *[x['col'] for x in self.enums]])
    self.enums.append({
      'values': values,
      'col': col
    })

  def get_enum(self, values: list) -> dict:
    """Get enum properties by name."""
    for x in self.enums:
      if x['values'] == values:
        return x
    raise KeyError(f"Enum with values {values} not found")

  def get_enum_range(self, values: list) -> str:
    """Get an enum's cell range."""
    x = self.get_enum(values)
    return helpers.column_to_range(
      col=x['col'],
      row=0,
      nrows=len(x['values']),
      fixed=True,
      sheet=self.enum_sheet
    )

  def get_column_code(self, table: str, column: str) -> str:
    """
    Get a column's code.

    Examples
    --------
    >>> layout = Layout()
    >>> layout.set_table('table', ['id', 'x'])
    >>> layout.get_column_code('table', 'id')
    'A'
    >>> layout.get_column_code('table', 'x')
    'B'
    """
    x = self.get_table(table)
    return helpers.column_index_to_code(x['columns'].index(column))

  def get_column_range(
    self,
    table: str,
    column: str,
    nrows: int = None,
    fixed: bool = False,
    absolute: bool = False,
    indirect: bool = False
  ) -> str:
    """
    Get a column's cell range.
    """
    x = self.get_table(table)
    col = x['columns'].index(column)
    cells = helpers.column_to_range(
      col=col,
      row=1,
      nrows=nrows or (self.max_rows - 1),
      fixed=fixed,
      sheet=x['sheet'] if absolute else None
    )
    if indirect:
      cells = f'INDIRECT({cells})'
    return cells

  # def build_column_dropdown(
  #   self,
  #   table: str,
  #   column: str,
  #   message: str = 'Value must be in list'
  # ) -> dict:
  #   cells = self.get_column_range(table=table, column=column)
  #   return {'values': cells, 'message': message}

  # def build_enum_dropdown(
  #   self,
  #   values: list,
  #   sheet: str,
  #   message: str = 'Value must be in list'
  # ) -> dict:
  #   cells = self.get_enum(values=values, sheet=sheet)['range']
  #   return {'values': cells, 'message': message}

  # def build_column_validations(
  #   self,
  #   table: str,
  #   column: str,
  #   dtype: str = None,
  #   required: bool = False,
  #   unique: bool = False,
  #   minimum: Union[int, float] = None,
  #   maximum: Union[int, float] = None,
  #   min_length: int = None,
  #   max_length: int = None,
  #   pattern: str = None
  # ) -> List[dict]:
  #   columns = self.get_table(table)['columns']
  #   defaults = {
  #     'col': self.get_column_code(table, column),
  #     'row': 2,
  #     'min_col': self.get_column_code(table, columns[0]),
  #     'max_col': self.get_column_code(table, columns[-1]),
  #     'max_row': f'${self.max_rows}',
  #     'ncols': len(columns)
  #   }
  #   checks = []

  #   # Field type
  #   if dtype in constants.TYPES:
  #     check = constants.TYPES[dtype]
  #     checks.append({
  #       'formula': check['valid'].format(**defaults),
  #       'message': check['message'],
  #       'ignore_blank': check['ignore_blank']
  #     })

  #   # Field constraints (except enum)
  #   constraints = {
  #       'required': required,
  #       'unique': unique,
  #       'minimum': minimum,
  #       'maximum': maximum,
  #       'min_length': min_length,
  #       'max_length': max_length,
  #       'pattern': pattern
  #   }
  #   for key, value in constraints.items():
  #       if value in (None, False, ''):
  #           continue
  #       check = constants.CONSTRAINTS[key]
  #       checks.append({
  #           'formula': check['valid'].format(**defaults, value=value),
  #           'message': check['message'].format(value=value),
  #           'ignore_blank': check['ignore_blank']
  #       })
  #   return checks

  def gather_column_checks(
    self,
    table: str,
    column: str,
    valid: bool,
    dtype: str = None,
    required: bool = False,
    unique: bool = False,
    minimum: Union[int, float] = None,
    maximum: Union[int, float] = None,
    min_length: int = None,
    max_length: int = None,
    pattern: str = None,
    enum: list = None,
    foreign_keys: List[Tuple[str, str]] = None
  ) -> List[constants.Check]:
    columns = self.get_table(table)['columns']
    defaults = {
      'col': self.get_column_code(table, column),
      'row': 2,
      'min_col': self.get_column_code(table, columns[0]),
      'max_col': self.get_column_code(table, columns[-1]),
      'max_row': f'${self.max_rows}',
      'ncols': len(columns)
    }
    f: Literal["valid", "invalid"] = "valid" if valid else "invalid"
    checks = []

    # Field type
    if dtype in constants.TYPES:
      check = constants.TYPES[dtype]
      checks.append({
        'formula': check[f].format(**defaults),
        'message': check['message'],
        'ignore_blank': check['ignore_blank']
      })

    # Field constraints (except enum)
    constraints = {
        'required': required,
        'unique': unique,
        'minimum': minimum,
        'maximum': maximum,
        'min_length': min_length,
        'max_length': max_length,
        'pattern': pattern
    }
    for key, value in constraints.items():
        if value in (None, False, ''):
            continue
        check = constants.CONSTRAINTS[key]
        checks.append({
            'formula': check[f].format(**defaults, value=value),
            'message': check['message'].format(value=value),
            'ignore_blank': check['ignore_blank']
        })

    # Range lookups
    check = constants.IN_RANGE
    # enum
    if enum:
      enum_range = self.get_enum_range(enum)
      checks.append({
        'formula': check[f].format(**defaults, range=enum_range),
        'message': check['message'].format(range=enum_range),
        'ignore_blank': check['ignore_blank']
      })
    # foreign keys
    for foreign_table, foreign_column in (foreign_keys or []):
      foreign_table = foreign_table or table
      column_range = self.get_column_range(
        foreign_table,
        foreign_column,
        absolute=foreign_table != table,
        fixed=True,
        indirect=self.indirect
      )
      checks.append({
          'formula': check[f].format(**defaults, range=column_range),
          'message': check['message'].format(range=column_range),
          'ignore_blank': check['ignore_blank']
      })
    return checks
