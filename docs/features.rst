Features and quirks
===================

Table name
----------

Frictionless Data limits ``resource.name`` (table name) to lowercase letters `a-z`,
numbers `0-9`, period `.`, dash `-`, and underscore `_`
(see `Data Resource <https://specs.frictionlessdata.io/data-resource/#name>`__).
``tablecloth`` has no such restrictions, so you can name your sheets however you wish.


Column type
-----------

``tablecloth`` supports a subset of the column types (``field.type``)
defined in the Frictionless Data
`Table Schema <https://specs.frictionlessdata.io/table-schema/#types-and-formats>`__.

.. list-table::
   :widths: 25 25 25
   :header-rows: 1

   * - Type
     - Micosoft Excel
     - Google Sheets
   * - ``string``
     - ✅
     - ✅
   * - ``integer``
     - ✅
     - ✅
   * - ``number``
     - ✅
     - ✅
   * - ``boolean``
     - ✅
     - ✅

.. note::

   **Date/time types** may be supported in the future.
   In Google Sheets, they can be simulated using a ``string`` with a
   ``pattern`` constraint (see `Column constraints`_ below).
   Micosoft Excel represents dates as numbers and does not support ``pattern``
   constraints, making implementation difficult.

Column constraints
------------------

``tablecloth`` supports all column constraints (``field.constraints``)
defined in the Frictionless Data
`Table Schema <https://specs.frictionlessdata.io/table-schema/#constraints>`__,
but not for all column types or all spreadsheet platforms.

.. list-table::
   :widths: 25 25 25
   :header-rows: 1

   * - Constraint
     - Micosoft Excel
     - Google Sheets
   * - ``required``
     - ✅
     - ✅
   * - ``unique``
     - ✅
     - ✅
   * - ``minLength`` (``min_length``)
     - ``string``
     - ``string``
   * - ``maxLength`` (``max_length``)
     - ``string``
     - ``string``
   * - ``minimum``
     - ``integer``, ``number``
     - ``integer``, ``number``
   * - ``maximum``
     - ``integer``, ``number``
     - ``integer``, ``number``
   * - ``pattern``
     - ❌
     - ``string``
   * - ``enum``
     - ``integer``, ``number``, ``string``.
       Unreliable behavior for strings starting with =, +, or '.
     - ``integer``, ``number``, ``string``.
       Strings starting with =, +, or ' should be prefixed with '.

Microsoft Excel does not have built-in regular expression support,
so there is no reliable way of implementing a ``pattern`` constraint.

Primary key
-----------

``tablecloth`` ignores any primary key defined in ``resource.schema.primaryKey``
(see `Table Schema <https://specs.frictionlessdata.io/table-schema/#primary-key>`__).
Instead, use the ``required`` and ``unique`` column constraints.
No fast and reliable way of validating a unique constraint across multiple columns
in spreadsheet software has yet been found.

Foreign keys
------------

``tablecloth`` supports foreign keys defined in ``resource.schema.foreignKeys``
(see `Table Schema <https://specs.frictionlessdata.io/table-schema/#foreign-keys>`__),
but treats composite foreign keys (involving multiple columns) as multiple
simple foreign keys (involving a single column).

Full validation of composite foreign keys may follow in the future.
