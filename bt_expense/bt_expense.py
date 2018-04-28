"""
bt_expense.py
~~~~~~~~~
Pull expenses from Excel Spreadsheet and upload to BigTime via REST HTTP call.
"""
import os
import json
from pprint import pprint as pp
# from pprint import pformat as pf
import requests as r
from openpyxl import load_workbook

# CD
# os.chdir('bt_expense')

# Constants
BASE = 'https://iq.bigtime.net/BigtimeData/api/v2'
UTF = 'utf-8'
# Global Variables
BT_LOOKUP = {'proj': {},
             'cat': {}, }


class Authorizer(object):
    """authorizes a BitTime REST API session.
    Can authorize using a user login and password or an API key.

    User login and password will be used to obtain an API key.
    If API key is provided, skip the step of obtaining API key"""

    def __init__(self, workbook_filename='Expenses.xlsx', staffsid=None):
        self.wb_name = workbook_filename
        self.staffsid = staffsid
        self.auth_header = self._build_credentials()
        self.userid = self.auth_header['userid']
        self.userpwd = self.auth_header['pwd']
        self.api_key = None
        self._authorized = False
        self.header = self.authorize_session()

    def _build_credentials(self):
        """Pulls Login information from the `Setup` worksheet. Return dictionary
        for Auth Header."""
        keys = get_values('Setup', 'A1', 'A4', workbook_name=self.wb_name)
        values = get_values('Setup', 'B1', 'B4', workbook_name=self.wb_name)
        header = {k: v for (k, v) in zip(keys, values)}
        header['Content-Type'] = 'application/json'
        return header

    def authorize_session(self):
        response = r.post('{}/session'.format(BASE),
                          headers={'Content-Type': 'application/json'},
                          data=json.dumps(self.auth_header).encode('utf-8'))
        if str(response.status_code)[0] is not '2':
            raise(ConnectionRefusedError)
        response_dict = json.loads(response.text)
        self.api_key = response_dict['token']
        self.staffsid = response_dict['staffsid']
        header = {'X-Auth-Token': self.api_key,
                  'X-Auth-Realm': self.auth_header['Firm'],
                  'Content-Type': self.auth_header['Content-Type']}
        self._authorized = True
        print('Session Header\n', header)
        return header


class Expensor(Authorizer):

    def post_expenses(self):
        projs = get_values('Expenses', 'F2', 'F121')
        cats = get_values('Expenses', 'G2', 'G121')
        dates = get_values('Expenses', 'C2', 'C121')
        costs = get_values('Expenses', 'D2', 'D121')
        notes = get_values('Expenses', 'E2', 'E121')

        expense_url = '{}/expense/detail'.format(BASE)
        for proj, cat, date, cost, note in zip(projs, cats,
                                               dates, costs, notes):
            # print(str(date)[:10])
            content = {'staffsid': int(self.staffsid),
                       'projectsid': int(proj),
                       'catsid': int(cat),
                       'dt': str(date)[:10],
                       'CostIN': cost,
                       'Nt': note,
                       'notes': 'March Expense',
                       'ApprovalStatus': 0}
            print(r.post(expense_url, headers=self.header,
                         data=json.dumps(content).encode()))
        return len(costs)


def get_wb(workbook_name='Expenses.xlsx'):
    return load_workbook(filename=workbook_name, data_only=True)


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
    sheet = get_wb(workbook_name)[sheet_name]
    if not stop:
        stop = sheet.max_row
    cells = [c[0].value for c in sheet[start:stop]]
    values = [c for c in cells if c is not None]
    return values


def get_picklist(auth_object, picklist_name):
    """Pulls a BigTime 'Picklist'
    Use to build project and expense catagory lookup tables.
    Requires Admin account."""
    # TODO: complete `get_picklist()` function
    valid_picklists = ['projects', 'ExpenseCodes']
    if picklist_name not in valid_picklists:
        raise ValueError('Not a valid picklist')
    pick_list_url = '{0}/picklist/{1}'.format(BASE, picklist_name)
    print(pick_list_url)
    response = r.get(pick_list_url, headers=auth_object.header)
    return response.json()
    # return response.json()


if __name__ == '__main__':
    print(__doc__)
    print('**DIR:', os.getcwd())
    build_lookup_dictn_from_excel()
    pp(BT_LOOKUP)
    # pp(build_credentials())
    NRC_AUTH = Authorizer()
    pp(NRC_AUTH.auth_header)
    pp(NRC_AUTH.api_key)
    print('*' * 79)
    # expense_codes = get_picklist(NRC_AUTH, 'ExpenseCodes')
    # with open('expense_codes.csv', 'w') as f_out:
    #     f_out.write('Id,Name')
    #     for expense_object in expense_codes:
    #         f_out.write('{},{}\n'.format(expense_object['Id'],
    #                                      expense_object['Name']))
    exp1 = Expensor(staffsid=859)
    pp(exp1.header)
    # pp(exp1.post_expenses())
