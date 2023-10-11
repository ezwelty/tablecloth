import re
import warnings
from typing import Any, List, Literal, Optional, Type

from . import constants, helpers


class Layout:
    """
    Tabular spreadsheet layout.

    Attributes
    ----------
    tables
        Tables added by :meth:`set_table`.
    enums
        Enums added by :meth:`set_enum`.
    enum_sheet
        Name of the sheet used to store enums.
    max_rows
        Maximum number of rows allowed per sheet.
        If `None`, column ranges are unbounded (e.g. `A2:A`).
        Otherwise, they are bounded (e.g. `A2:A1000`).
    max_name_length
        Maximum length of sheet names (in characters).
    """

    def __init__(
        self,
        enum_sheet: str = 'lists',
        max_rows: int = None,
        max_name_length: int = None,
    ) -> None:
        self.tables: List[constants.Table] = []
        self.enums: List[constants.Enum] = []
        self.enum_sheet = enum_sheet
        self.max_rows = max_rows
        self.max_name_length = max_name_length

    @classmethod
    def from_package(cls: Type['Layout'], package: dict, **kwargs: Any) -> 'Layout':
        """
        Initialize a layout from a Frictionless Data Tabular Data Package.

        Sets all tables and enums found in the package using default settings.

        Parameters
        ----------
        package
            Frictionless Data Tabular Data Package specification.
            See https://specs.frictionlessdata.io/tabular-data-package.
        **kwargs
            Arguments to :class:`Layout`.
        """
        layout = cls(**kwargs)
        for resource in package['resources']:
            layout.set_table(
                table=resource['name'],
                columns=[field['name'] for field in resource['schema']['fields']],
            )
            for field in resource['schema']['fields']:
                values = field.get('constraints', {}).get('enum')
                if values:
                    layout.set_enum(values)
        return layout

    def set_table(self, table: str, columns: List[str], sheet: str = None) -> None:
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
            raise ValueError(f'Sheet name cannot be longer than {self.max_name_length}')
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
        formulas = [x for x in values if isinstance(x, str) and re.match(r"^[+=']", x)]
        if formulas:
            warnings.warn(
                f"Enum values start with +, =, or ' {formulas}. "
                'Expect unexpected behavior.'
            )
        if values in [x['values'] for x in self.enums]:
            return
        if self.enums:
            col = max([x['col'] for x in self.enums]) + 1
        else:
            col = 0
        self.enums.append({'values': values, 'col': col})

    def get_enum(self, values: list) -> dict:
        """Get enum properties by name."""
        for x in self.enums:
            if x['values'] == values:
                return x
        raise KeyError(f'Enum with values {values} not found')

    def get_enum_range(self, values: list, indirect: bool = False) -> str:
        """
        Get an enum's cell range.

        Parameters
        ----------
        values
            Enum values.
        indirect
            Whether to wrap range in `INDIRECT` function.
            See https://support.google.com/docs/answer/3093377.
        """
        x = self.get_enum(values)
        return helpers.column_to_range(
            col=x['col'],
            row=0,
            nrows=len(x['values']),
            fixed=True,
            sheet=self.enum_sheet,
            indirect=indirect,
        )

    def get_column_code(self, table: str, column: str) -> str:
        """
        Get a column's code.

        Parameters
        ----------
        table
            Table name.
        column
            Column name.

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
        index = x['columns'].index(column)
        return helpers.column_index_to_code(index)

    def get_column_range(
        self,
        table: str,
        column: str,
        nrows: int = None,
        absolute: bool = False,
        fixed: bool = False,
        indirect: bool = False,
    ) -> str:
        """
        Get a column's cell range (not including the header).

        Parameters
        ----------
        table
            Table name.
        column
            Column name.
        nrows
            Number of rows to include.
            If None, includes all rows (to `self.max_rows` or unbounded if undefined).
        absolute
            Whether to refer to the range by sheet name (e.g. 'Sheet1'!A2:A).
        fixed
            Whether to use a fixed range (e.g. $A$2:$A).
        indirect
            Whether to wrap range in `INDIRECT` function.
            See https://support.google.com/docs/answer/3093377.
        """
        x = self.get_table(table)
        col = x['columns'].index(column)
        return helpers.column_to_range(
            col=col,
            row=1,
            nrows=nrows or (self.max_rows - 1 if self.max_rows else None),
            sheet=x['sheet'] if absolute else None,
            fixed=fixed,
            indirect=indirect,
        )

    def select_column_dropdown(
        self,
        table: str,
        column: str,
        dtype: str = None,
        constraints: constants.Constraints = None,
        foreign_keys: List[constants.ForeignKey] = None,
        indirect: bool = False,
    ) -> Optional[constants.Dropdown]:
        """
        Whether and what to use as a dropdown for a column's data validation.

        The dropdown (if any) is selected in the following order:

        * Boolean data type (`dtype: 'boolean'`). Uses values (TRUE, FALSE).
        * Enumerated value constraint (`constraints.enum`).
        * Foreign key. If multiple `foreign_keys` match the column, the first is used.
          A foreign key involving multiple columns is reduced to a simple foreign key
          between `column` and the corresponding foreign column.

        Parameters
        ----------
        table
            Table name.
        column
            Column name.
        dtype
            Column type. See
            https://specs.frictionlessdata.io/table-schema/#types-and-formats.
        constraints
            Column constraints. See
            https://specs.frictionlessdata.io/table-schema/#constraints.
            An `enum` constraint must already be registered with :meth:`set_enum`.
        foreign_keys
            Foreign key constraints. See
            https://specs.frictionlessdata.io/table-schema/#foreign-keys.
            The referenced table(s) and column(s) must already be registered with
            :meth:`set_table`.
        indirect
            Whether to wrap cell ranges (for enum constraints and foreign keys)
            in the `INDIRECT` function.
            See https://support.google.com/docs/answer/3093377.

        Returns
        -------
        dropdown :
            Either None or a dictionary with the following keys:

            * source: Either 'boolean', 'foreign_key', or 'enum'.
            * values: Either a list (for boolean) or a cell range.
        """
        # Boolean
        if dtype == 'boolean':
            return {'source': 'boolean', 'values': ['TRUE', 'FALSE']}
        # Foreign key
        keys = helpers.reduce_foreign_keys(
            foreign_keys or [], table=table, column=column
        )
        if keys:
            values = self.get_column_range(
                table=keys[0][0] or table,
                column=keys[0][1],
                fixed=True,
                absolute=keys[0][0] is not None,
                indirect=indirect,
            )
            return {'source': 'foreign_key', 'values': values}
        # Enum
        enum = (constraints or {}).get('enum')
        if enum:
            values = self.get_enum_range(enum, indirect=indirect)
            return {'source': 'enum', 'values': values}
        return None

    def gather_column_checks(
        self,
        table: str,
        column: str,
        valid: bool,
        dtype: str = None,
        constraints: constants.Constraints = None,
        foreign_keys: List[constants.ForeignKey] = None,
        indirect: bool = False,
    ) -> List[constants.Check]:
        """
        Gather all checks for a column.

        Parameters
        ----------
        table
            Table name.
        column
            Column name.
        valid
            Whether formulas should return TRUE for valid (True) or invalid (False)
            values.
        dtype
            Column type. See
            https://specs.frictionlessdata.io/table-schema/#types-and-formats
            (and `constants.TYPES` for which are supported).
        constraints
            Column constraints. See
            https://specs.frictionlessdata.io/table-schema/#constraints.
            Both snake_case and camelCase are supported.
            An `enum` constraint must first be registered with `self.set_enum`.
        foreign_keys
            Foreign key constraints. See
            https://specs.frictionlessdata.io/table-schema/#foreign-keys.
            Composite foreign keys (with multiple columns) cannot be readily enforced
            in a spreadsheet, so they are treated as multiple simple foreign keys.
        indirect
            Whether to wrap cell ranges (for enum constraints and foreign keys)
            in the `INDIRECT` function.
            See https://support.google.com/docs/answer/3093377.
        """
        columns = self.get_table(table)['columns']
        defaults = {
            'col': self.get_column_code(table, column),
            'row': 2,
            'min_col': self.get_column_code(table, columns[0]),
            'max_col': self.get_column_code(table, columns[-1]),
            'max_row': f'${self.max_rows}' if self.max_rows else '',
            'ncols': len(columns),
        }
        f: Literal['valid', 'invalid'] = 'valid' if valid else 'invalid'
        checks = []

        # Field type
        if dtype in constants.TYPES:
            check = constants.TYPES[dtype]
            checks.append(
                {
                    'formula': check[f].format(**defaults),
                    'message': check['message'],
                    'ignore_blank': check['ignore_blank'],
                }
            )

        # Format constraints
        constraints = {
            helpers.camel_to_snake_case(key): value
            for key, value in (constraints or {}).items()
            if value not in (None, False, '', [])
        }

        # Field constraints (except enum)
        for key, value in constraints.items():
            if key not in constants.CONSTRAINTS:
                continue
            check = constants.CONSTRAINTS[key]
            checks.append(
                {
                    'formula': check[f].format(**defaults, value=value),
                    'message': check['message'].format(value=value),
                    'ignore_blank': check['ignore_blank'],
                }
            )

        # Range lookups
        check = constants.IN_RANGE
        # enum
        if 'enum' in constraints:
            enum_range = self.get_enum_range(constraints['enum'], indirect=indirect)
            checks.append(
                {
                    'formula': check[f].format(**defaults, range=enum_range),
                    'message': check['message'].format(range=enum_range),
                    'ignore_blank': check['ignore_blank'],
                }
            )
        # foreign keys
        for key in foreign_keys or []:
            columns = helpers.to_list(key['fields'])
            foreign_columns = helpers.to_list(key['reference']['fields'])
            try:
                i = columns.index(column)
            except ValueError:
                continue
            foreign_column = foreign_columns[i]
            foreign_table = key['reference'].get('resource') or table
            column_range = self.get_column_range(
                foreign_table,
                foreign_column,
                absolute=foreign_table != table,
                fixed=True,
                indirect=indirect,
            )
            checks.append(
                {
                    'formula': check[f].format(**defaults, range=column_range),
                    'message': check['message'].format(range=column_range),
                    'ignore_blank': check['ignore_blank'],
                }
            )
        return checks
