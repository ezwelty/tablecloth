`tablecloth`
============

[![codecov](https://codecov.io/gh/ezwelty/tablecloth/branch/main/graph/badge.svg?token=RP9E7WLFCI)](https://codecov.io/gh/ezwelty/tablecloth)
![tests](https://github.com/ezwelty/tablecloth/actions/workflows/tests.yaml/badge.svg)

Generate spreadsheet templates for data entry based on metadata.

## Installation

```sh
pip install "tablecloth[excel,gsheets] @ git+https://github.com/ezwelty/tablecloth"
# Or with only one output format supported:
# pip install "tablecloth[excel] @ git+https://github.com/ezwelty/tablecloth"
# pip install "tablecloth[gsheets] @ git+https://github.com/ezwelty/tablecloth"
```

* `[excel]`: Adds (optional) support for Microsoft Excel
* `[gsheets]`: Adds (optional) support for Google Sheets

## Example

```py
import tablecloth.excel

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

## Development

Clone the repository and use `poetry` to install `tablecloth` into a virtual environment
linked to your current Python version:

```bash
git clone https://github.com/ezwelty/tablecloth
cd tablecloth
poetry install
```

To avoid committing code that breaks tests or violates the style guide,
consider installing [`pre-commit`](https://pre-commit.com) (if needed)
and installing the hooks:

```bash
pre-commit install
```

Useful commands are listed in [`.pre-commit-config.yaml`](pre-commit-config.yaml).
For example, to run most tests:

```bash
poetry run pytest --doctest-modules src tests
```

### Google Sheets

To run tests for Google Sheets, you will need to create a Google Cloud project,
add a service account, and create and download a JSON key
(see https://pygsheets.readthedocs.io/en/stable/authorization.html#service-account).
Then, copy the example environment file:

```bash
cp .env.example .env
```

and fill in the content of the JSON key.
Optionally, add your personal Google account email so that Google Sheets created by your
service account are shared with you.
Once configured, use the `--gsheets` flag include Google Sheets tests:

```bash
poetry run pytest --doctest-modules src tests --gsheets
```
