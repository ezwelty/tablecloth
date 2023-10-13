"""Lay tables out onto a spreadsheet."""
from __future__ import annotations

import re
import warnings
from typing import Any, List, Literal, Type, cast

from . import constants, helpers


class Layout:
    """
    Tabular spreadsheet layout.

    Parameters
    ----------
    enum_sheet
        Name of the sheet used to store enums.
    max_rows
        Maximum number of rows allowed per sheet.
        If `None`, column ranges are unbounded (e.g. `A2:A`).
        Otherwise, they are bounded (e.g. `A2:A1000`).
    max_name_length
        Maximum length of sheet names.

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
    max_name_length
        Maximum length of sheet names.

    Examples
    --------
    Add tables and enums to the layout.

    >>> layout = Layout()
    >>> layout.set_table('tree', ['id', 'name', 'type'])
    >>> layout.set_table('branch', ['tree_id', 'attached'])
    >>> tree_types = ['deciduous', 'evergreen']
    >>> layout.set_enum(tree_types)

    Get table and enum properties.

    >>> layout.get_table('tree')
    {'sheet': 'tree', 'table': 'tree', 'columns': ['id', 'name', 'type']}
    >>> layout.get_enum(tree_types)
    {'values': ['deciduous', 'evergreen'], 'col': 0}

    Get column and enum cell ranges.

    >>> layout.get_column_range('tree', 'id', absolute=True, fixed=True)
    "'tree'!$A$2:$A"
    >>> layout.get_column_range('branch', 'tree_id')
    'A2:A'
    >>> layout.get_enum_range(tree_types)
    "'lists'!$A$1:$A$2"

    Determine which dropdown (if any) to use for a column.

    >>> layout.select_column_dropdown('tree', 'type', constraints={'enum': tree_types})
    {'source': 'enum', 'options': "'lists'!$A$1:$A$2"}
    >>> layout.select_column_dropdown('branch', 'attached', dtype='boolean')
    {'source': 'boolean', 'options': ['TRUE', 'FALSE']}
    >>> foreign_keys = [{
    ...    'fields': ['tree_id'],
    ...     'reference': {'resource': 'tree', 'fields': ['id']},
    ... }]
    >>> layout.select_column_dropdown('branch', 'tree_id', foreign_keys=foreign_keys)
    {'source': 'foreign_key', 'options': "'tree'!$A$2:$A"}
    >>> layout.select_column_dropdown('tree', 'id') is None
    True

    Gather all checks for a column.

    >>> checks = layout.gather_column_checks(
    ...     'branch', 'tree_id', valid=True, dtype='integer', foreign_keys=foreign_keys
    ... )
    >>> import pprint
    >>> pprint.pprint(checks)
    [{'formula': 'IF(ISNUMBER(A2), INT(A2) = A2, FALSE)',
      'ignore_blank': True,
      'message': 'integer'},
     {'formula': "ISNUMBER(MATCH(A2, 'tree'!$A$2:$A, 0))",
      'ignore_blank': True,
      'message': "in the range 'tree'!$A$2:$A"}]
    """

    def __init__(
        self,
        enum_sheet: str = 'lists',
        max_rows: int | None = None,
        max_name_length: int | None = None,
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

    def set_table(
        self, table: str, columns: List[str], sheet: str | None = None
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

        Raises
        ------
        ValueError
            Table `table` already exists.
        ValueError
            Sheet `sheet` already exists.
        ValueError
            Sheet name cannot be longer than :attr:`max_name_length`.
        ValueError
            Columns `columns` are not unique.
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

    def get_table(self, table: str) -> constants.Table:
        """Get table properties by name."""
        for x in self.tables:
            if x['table'] == table:
                return x
        raise KeyError(f"Table '{table}' not found")

    def set_enum(self, values: list) -> None:
        """
        Add an enum (if new) to the enum sheet.

        Parameters
        ----------
        values
            Enum values. Should be unique (since used in dropdown) and not start with
            special spreadsheet characters (+, =, '), but this is not required.

        Raises
        ------
        ValueError
            Enum values (`values`) has zero length.
        """
        if not len(values):
            raise ValueError('Enum values (`values`) have zero length')
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

    def get_enum(self, values: list) -> constants.Enum:
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
            Whether to wrap range in INDIRECT function.
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
        """Get a column's code."""
        x = self.get_table(table)
        index = x['columns'].index(column)
        return helpers.column_index_to_code(index)

    def get_column_range(
        self,
        table: str,
        column: str,
        nrows: int | None = None,
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
            Whether to wrap range in INDIRECT function.
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
        dtype: str | None = None,
        constraints: constants.Constraints | None = None,
        foreign_keys: List[constants.ForeignKey] | None = None,
        indirect: bool = False,
    ) -> constants.Dropdown | None:
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
            in the INDIRECT function.
            See https://support.google.com/docs/answer/3093377.
        """
        # Boolean
        if dtype == 'boolean':
            return {'source': 'boolean', 'options': ['TRUE', 'FALSE']}
        # Foreign key
        keys = helpers.reduce_foreign_keys(
            foreign_keys or [], table=table, column=column
        )
        if keys:
            options = self.get_column_range(
                table=keys[0][0] or table,
                column=keys[0][1],
                fixed=True,
                absolute=keys[0][0] is not None,
                indirect=indirect,
            )
            return {'source': 'foreign_key', 'options': options}
        # Enum
        enum = (constraints or {}).get('enum')
        if enum:
            options = self.get_enum_range(enum, indirect=indirect)
            return {'source': 'enum', 'options': options}
        return None

    def gather_column_checks(
        self,
        table: str,
        column: str,
        valid: bool,
        dtype: str | None = None,
        constraints: constants.Constraints | None = None,
        foreign_keys: List[constants.ForeignKey] | None = None,
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
            An `enum` constraint must already be registered with :meth:`set_enum`.
        foreign_keys
            Foreign key constraints. See
            https://specs.frictionlessdata.io/table-schema/#foreign-keys.
            Composite foreign keys (with multiple columns) cannot be readily enforced
            in a spreadsheet, so they are treated as multiple simple foreign keys.
        indirect
            Whether to wrap cell ranges (for enum constraints and foreign keys)
            in the INDIRECT function.
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
            template = constants.TYPES[dtype]
            check: constants.Check = {
                'formula': template[f].format(**defaults),
                'message': template['message'],
                'ignore_blank': template['ignore_blank'],
            }
            checks.append(check)

        # Format constraints
        clean_constraints = cast(
            constants.Constraints,
            {
                helpers.camel_to_snake_case(key): value
                for key, value in (constraints or {}).items()
                if value not in (None, False, '', [])
            },
        )

        # Field constraints (except enum)
        for key, value in clean_constraints.items():
            if key not in constants.CONSTRAINTS:
                continue
            template = constants.CONSTRAINTS[key]
            check = {
                'formula': template[f].format(**defaults, value=value),
                'message': template['message'].format(value=value),
                'ignore_blank': template['ignore_blank'],
            }
            checks.append(check)

        # Range lookups
        template = constants.IN_RANGE
        # enum
        if 'enum' in clean_constraints:
            enum_range = self.get_enum_range(
                values=clean_constraints['enum'], indirect=indirect
            )
            check = {
                'formula': template[f].format(**defaults, range=enum_range),
                'message': template['message'].format(range=enum_range),
                'ignore_blank': template['ignore_blank'],
            }
            checks.append(check)
        # foreign keys
        simple_foreign_keys = helpers.reduce_foreign_keys(
            foreign_keys or [], table=table, column=column
        )
        for foreign_table, foreign_column in simple_foreign_keys:
            column_range = self.get_column_range(
                table=foreign_table or table,
                column=foreign_column,
                absolute=foreign_table is not None,
                fixed=True,
                indirect=indirect,
            )
            check = {
                'formula': template[f].format(**defaults, range=column_range),
                'message': template['message'].format(range=column_range),
                'ignore_blank': template['ignore_blank'],
            }
            checks.append(check)
        return checks
