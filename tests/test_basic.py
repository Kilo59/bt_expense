"""
basic_tests.py
~~~~~~~~~~~~~~
simple tests for bt_expense.
"""
import os
import unittest

# import pytest

from context import bt_expense as bte
from context import fixpath

TEST_DIR = fixpath(os.path.abspath(os.path.dirname(__file__)))
ROOT_DIR = fixpath(os.path.dirname(TEST_DIR))
MAIN_DIR = fixpath('{}/bt_expense'.format(ROOT_DIR))


class SmokeTest(unittest.TestCase):
    """Test that nothing is on fire."""
    def setUp(self):
        os.chdir(TEST_DIR)

    def tearDown(self):
        os.chdir(ROOT_DIR)

    def test_pulling_column_values(self):
        os.chdir(MAIN_DIR)
        a1 = bte.get_values('Expenses', 'A1')[0]
        self.assertEqual(a1, 'Project')


if __name__ == "__main__":
    print(__doc__)
    print(__file__)
    print('root:', ROOT_DIR)
    print('test:', TEST_DIR)
    print('main:', MAIN_DIR)
    unittest.main()
