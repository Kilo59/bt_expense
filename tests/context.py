# -*- coding: utf-8 -*-
"""
context.py
~~~~~~~~~~
Access main module from tests folder
"""
import os
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../bt_expense')))

import bt_expense

print('USING context.py')

if __name__ == '__main__':
    print(bt_expense.__doc__)
