"""
bt_expense.py
~~~~~~~~~
Pull expenses from Excel Spreadsheet and upload to BigTime via REST HTTP call.
"""
import os
from pprint import pprint as pp
# from pprint import pformat as pf
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


class Authorizer(object):
    """Autherizes a BitTime REST API session.
    Can authorize using a user login and password or an API key.

    User login and password will be used to obtain an API key.
    If API key is provided, skip the step of obtaining API key"""

    def __init__(self, workbook_filename='Expenses.xlsx'):
        self.wb_name = workbook_filename
        self.header = self._build_credentials()
        self.userid = self.header['User Id']
        self.userpwd = self.header['Password']
        self.api_key = self._check_user_provided_key()
        self._authorized = False

    def _build_credentials(self):
        """Pulls Login information from the `Setup` worksheet. Return dictionary
        for Auth Header."""
        keys = get_values('Setup', 'A1', 'A4', workbook_name=self.wb_name)
        values = get_values('Setup', 'B1', 'B4', workbook_name=self.wb_name)
        header = {k: v for (k, v) in zip(keys, values)}
        # TODO: Format for BigTime
        return header

    def _check_user_provided_key(self):
        """Checks the Excel workbook API key value.
        If the length of the API key matches the expected length of the API
        key, assume the key is valid."""
        api_key_value = get_values('Setup', 'B5', 'B5',
                                   workbook_name=self.wb_name)[0]
        # TODO: change to length of of API key
        if len(api_key_value) <= 15:
            print('\tNo API key')
            return None
        else:
            print('\tAPI key provided')
            return api_key_value


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


def get_values(sheet_name, start, stop=None, workbook_name='Expenses.xlsx'):
    """Pulls a column (or section) of values from a Worksheet.
    Returns a list."""
    values = []
    sheet = get_wb()[sheet_name]
    if not stop:
        stop = sheet.max_row
    cells = [c[0].value for c in sheet[start:stop]]
    values = [c for c in cells if c is not None]
    return values


def get_picklist(picklist_name):
    """Pulls a BigTime 'Picklist'
    Use to build project and expense catagory lookup tables"""
    # TODO: complete `get_picklist()` function
    valid_picklists = ['projects', 'ExpenseCodes']
    if picklist_name not in valid_picklists:
        raise ValueError('Not a valid picklist')
    # header = build_credentials()
    return picklist_name


if __name__ == '__main__':
    print(__doc__)
    print('**DIR:', os.getcwd())
    build_lookup_dictn_from_excel()
    pp(BT_LOOKUP)
    # pp(build_credentials())
    NRC_AUTH = Authorizer()
    pp(NRC_AUTH.header)
    pp(NRC_AUTH.api_key)
