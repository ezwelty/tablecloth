tablecloth
==========

.. toctree::
   :hidden:
   :maxdepth: 1

   self
   reference
   license

Generate spreadsheet templates for data entry based on a data schema!


Installation
------------

To install the Hypermodern Python project,
run this command in your terminal:

.. code-block:: console

   $ pip install tablecloth[excel,gsheets]

where ``excel`` adds optional support for Microsoft Excel,
and ``gsheets`` adds optional support for Google Sheets.

Usage
-----

For example, to create a Microsoft Excel template, ...

.. code-block:: python

   import tablecloth.excel

   # Describe the data following the Frictionless Tabular Data Package
   package = {
      'resources': [{
         'name': 'tree',
         'schema': {
         'fields': [{
            'name': 'id',
            'type': 'integer',
            'constraints': {'required': True, 'unique': True},
         }, {
            'name': 'type',
            'type': 'string',
            'constraints': {'enum': ['deciduous', 'evergreen']},
         }]
         }
      }, {
         'name': 'branch',
         'schema': {
         'foreignKeys': [{
            'fields': ['tree_id'],
            'reference': {'resource': 'tree', 'fields': ['id']},
         }],
         'fields': [{
            'name': 'tree_id',
            'type': 'integer',
         }, {
            'name': 'height',
            'type': 'number',
            'constraints': {'minimum': 0}
         }]
         }
      }]
   }

   # Build the template
   tablecloth.excel.write_template(package, path='template.xlsx')
