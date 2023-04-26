`tablecloth`
============

[![codecov](https://codecov.io/gh/ezwelty/tablecloth/branch/main/graph/badge.svg?token=RP9E7WLFCI)](https://codecov.io/gh/ezwelty/tablecloth)
![tests](https://github.com/ezwelty/tablecloth/actions/workflows/tests.yaml/badge.svg)

Generate spreadsheet templates for data entry based on metadata.

## Installation

```sh
pip install "tablecloth[excel] @ git+https://github.com/ezwelty/tablecloth"
```

The `[excel]` adds (optional) support for Microsoft Excel templates.
Optional support for Google Sheets should follow shortly.

## Example

```py
import tablecloth

# Load a Frictionless Tabular Data Package specification
package = {
  'name': 'package',
  'resources': [{
    'name': 'main',
    'schema': {
      'fields': [{
        'name': 'id',
        'description': 'Any positive integer',
        'type': 'integer',
        'constraints': {'required': True, 'unique': True, 'minimum': 1},
      }, {
        'name': 'number_minmax',
        'description': 'A number in interval [1, 10]',
        'type': 'number',
        'constraints': {'minimum': 1, 'maximum': 10},
      }, {
        'name': 'boolean',
        'description': 'Any boolean',
        'type': 'boolean',
      }, {
        'name': 'string_minmax',
        'description': 'Any string with length between 2 and 3',
        'type': 'string',
        'constraints': {'minLength': 2, 'maxLength': 3},
      }, {
        'name': 'string_enum',
        'description': 'One of the first three letters of the alphabet',
        'type': 'string',
        'constraints': {'enum': ['a', 'b', 'c']},
      }]
    }
  }, {
    'name': 'secondary',
    'schema': {
      'foreignKeys': [{
        'fields': ['main_id'],
        'reference': {'resource': 'main', 'fields': ['id']},
      }],
      'fields': [{
        'name': 'main_id',
        'description': 'Any value in main.id',
        'type': 'integer',
      }]
    }
  }]
}

# Use field descriptions as header comments
comments = {
  resource['name']: [
    field['description'] for field in resource['schema']['fields']
  ] for resource in package['resources']
}

# Build Excel template
tablecloth.excel.write_template(
  package, path='template.xlsx', header_comments=comments
)
```
