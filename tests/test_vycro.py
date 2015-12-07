###############################################################################
#
# Module: tests.py
# Author: Zach Gannon
# Revise: 2015-12-01
#
###############################################################################

from __future__ import absolute_import, print_function, division
import os
from vycro import MacroWrapper

HERE = os.path.dirname(os.path.abspath(__file__))
WORKBOOK = 'TestWorkbook.xlsm'
PATH = os.path.join(HERE, WORKBOOK)
MACRO = 'TestMacro'
WORKBOOK_KWARGS = {
    "ReadOnly": "True",
}


def test_open_workbook_context():
    mw = MacroWrapper()
    with mw.open_workbook(workbook=PATH, **WORKBOOK_KWARGS) as wb:
        assert wb is not None


def test_close_workbook():
    mw = MacroWrapper()
    mw.open_workbook(workbook=PATH, **WORKBOOK_KWARGS)
    assert mw.close_workbook() is True


def test_run_macro():
    mw = MacroWrapper()
    with mw.open_workbook(workbook=PATH, **WORKBOOK_KWARGS):
        assert mw.run_macro(MACRO) is True
