"""Sphinx configuration."""
import datetime

project = 'tablecloth'
author = 'Ethan Welty'
copyright = f'{datetime.date.today().year}, {author}'
extensions = [
    'sphinx.ext.intersphinx',
    'sphinx.ext.autodoc',
    'sphinx.ext.napoleon',
    'sphinx_autodoc_typehints',
]
html_theme = 'sphinx_rtd_theme'
autodoc_member_order = 'bysource'
intersphinx_mapping = {
    'python': ('https://docs.python.org/3', None),
    'xlsxwriter': ('https://xlsxwriter.readthedocs.io', None),
    'pygsheets': ('https://pygsheets.readthedocs.io/en/stable', None),
}
