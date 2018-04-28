"""
basic_tests.py
~~~~~~~~~~~~~~
simple tests for bt_expense.
"""
import os
import unittest
import pytest
from sys import version_info
# import pytest

from context import bt_expense as bte
from context import fixpath

PYTHON_VER = version_info[0]

TEST_DIR = fixpath(os.path.abspath(os.path.dirname(__file__)))
ROOT_DIR = fixpath(os.path.dirname(TEST_DIR))
MAIN_DIR = fixpath('{}/bt_expense'.format(ROOT_DIR))


def test_pulling_column_values():
    try:
        os.chdir(MAIN_DIR)
    except FileNotFoundError:
        pass
    a1 = bte.get_values('Expenses', 'A1',
                        workbook_name='Expenses_Template.xlsx')[0]
    assert a1, 'Project'
    os.chdir(ROOT_DIR)


class AuthorizerTests(unittest.TestCase):
    """Tests related to the Authorizer Class"""
    def setUp(self):
        os.chdir(MAIN_DIR)
        print('SetUp')

    def tearDown(self):
        print('tearDown')
        os.chdir(ROOT_DIR)

    def test_authorizer_object_creation(self):
        bte.Authorizer('bt_expense/Expenses_Template.xlsx')


if __name__ == "__main__":
    print('Python {}'.format(PYTHON_VER))
    print(__doc__)
    print(__file__)
    print('root:', ROOT_DIR)
    print('test:', TEST_DIR)
    print('main:', MAIN_DIR)
    pytest.main(args=['-v'])
