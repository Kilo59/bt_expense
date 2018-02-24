"""
bt_expense.py
~~~~~~~~~
Pull expenses from Excel Spreadsheet and upload to BigTime via REST HTTP call.
"""
import os
from pprint import pprint as pp
# import openpyxl as opxl
# import requests as r
from openpyxl import load_workbook

# CD
os.chdir('bt_expense')

# Constants
BASE = 'https://iq.bigtime.net/BigtimeData/api/v2'
UTF = 'utf-8'
# Global Variables
BT_LOOKUP = {'proj': {},
             'cat': {}, }


def get_wb(workbook_name='Expenses.xlsx'):
    return load_workbook(filename=workbook_name)


def build_lookup_dictn_from_excel():
    """Build lookup_dictionaries from the excel workbook"""
    project_names = get_values('Projects', 'A2')
    project_ids = get_values('Projects', 'B2')
    category_names = get_values('Categories', 'A2')
    category_ids = get_values('Categories', 'B2')
    BT_LOOKUP['proj'] = dict(zip(project_ids, project_names))
    BT_LOOKUP['cat'] = dict(zip(category_ids, category_names))
    return project_ids, category_ids


def get_values(sheet_name, start, stop=None):
    """Pulls a column (or section) of values from a Worksheet.
    Returns a list."""
    values = []
    sheet = get_wb()[sheet_name]
    if not stop:
        stop = sheet.max_row
    cells = [c[0].value for c in sheet[start:stop]]
    values = [c for c in cells if c is not None]
    return values


def build_credentials():
    """Pulls Login information from the `Setup` worksheet. Return dictionary
    for Auth Header."""
    keys = get_values('Setup', 'A1', 'A4')
    values = get_values('Setup', 'B1', 'B4')
    header = {k: v for (k, v) in zip(keys, values)}
    # TODO: Format for BigTime
    return header


def get_picklist(picklist_name):
    """Pulls a BigTime 'Picklist'
    Use to build project and expense catagory lookup tables"""
    # TODO: complete `get_picklist()` function
    valid_picklists = ['projects', 'ExpenseCodes']
    if picklist_name not in valid_picklists:
        raise ValueError('Not a valid picklist')
    header = build_credentials()
    return picklist_name


if __name__ == '__main__':
    print(__doc__)
    print('**DIR:', os.getcwd())
    build_lookup_dictn_from_excel()
    pp(BT_LOOKUP)
    pp(build_credentials())
