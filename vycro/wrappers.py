###############################################################################
#
# Module: wrappers.py
# Author: Zachary Gannon
# Revise: 2015-12-01
# Python: 2.7.X, 3.X
#
# Requirements:
#   pywin32 -- http://sourceforge.net/projects/pywin32/
#
# Description:
#   Vycro is wrapper for the pywin32 module, to allow ease of use with Excel
#   workbooks and VBA macros embedded in such. This wrapper offers users
#   a robust, yet simple API with examples galore.
#
###############################################################################

"""Module provides ease of use when calling VBA macros embedded in Excel."""

# Standard Python module dependencies
from __future__ import absolute_import, division, print_function
from contextlib import contextmanager
import os

# Third-Party Python module dependencies
import win32com.client as win32  # from pywin32


class MacroWrapper(object):
    """Class defines methods to easily manipulate excel workbooks and
    execute VBA macros.

    NOTE: All function parameters must be in DOUBLE QUOTES. That includes:
        Strings, Chars, Integers, Floats, Booleans
    This is due to VBA using single quotes as comment openers.

    """

    def __init__(self, parent=None):
        self.here = os.path.dirname(os.path.abspath(__file__))
        self.parent = parent
        self.excel_app = None

    @contextmanager
    def open_workbook(self, workbook, **workbook_kwargs):
        """Open a MS Excel workbook and initialize Excel object.
        Use the <with> keyword to open and close the workbook implicitly.

        Usage:
            mw = vycro.wrappers.MacroWrapper()
            wb_kwargs = {"ReadOnly":"True"}
            with mw.open_workbook("/full/path/with.extension", **wb_kwargs):
                mw.run_macro("macro_name")  # See declaration for function args

        Keyword arguments:
            workbook (String): Full path and name of workbook to open
            workbook_kwargs (Dictionary): Standard VBA Excel obtions passed
                when opening Excel workbook objects

        """
        import pythoncom
        pythoncom.CoInitialize()
        self.excel_app = win32.Dispatch("Excel.Application")
        try:
            self.excel_app.Workbooks.Open(Filename=workbook, **workbook_kwargs)
            yield self.excel_app
        except Exception as err:
            print('\nERROR in open_workbook function:\n\t{}'.format(str(err)))
        finally:
            self.close_workbook()

    def close_workbook(self):
        """Closes a Microsoft Excel workbook and deletes associated object.

        NOTE: If the <with> keyword was used to open the workbook, there is no
        need to call the close_workbook function.

        Usage:
            mw = vycro.wrappers.MacroWrapper()
            mw.open_workbook("/full/path/with.extension")
            mw.run_macro("macro_name")  # See declaration for function args
            mw.close_workbook()

        """
        try:
            self.excel_app.Workbooks(1).Close(False)
            self.excel_app.Application.Quit()
            self.excel_app = None
            return True
        except AttributeError:
            print(
                'WARNING: It appears as though the workbook has already been '
                'properly closed.'
            )
            return True
        except Exception as err:
            print(
                '\nERROR in close_workbook function:\n\t{}'
                .format(str(err))
            )
            return False

    def run_macro(self, macro_name, *macro_args):
        """Executes specified macro for given arguments.

        Usage:
            mw = vycro.wrappers.MacroWrapper()
            macro_name = "Macro1"
            macro_args = ["True", "1"]
            with mw.open_workbook("/full/path/with.extension"):
                mw.run_macro(macro_name, *macro_args)

        Keyword arguments:
            macro_name (String): Macro to execute
            macro_args (List): Arguments to send to macro

        """

        if self.excel_app:
            try:
                self.excel_app.Run(macro_name, *macro_args)
                return True
            except Exception as err:
                print('\nERROR in run_macro function:\n\t{}'.format(str(err)))
                return False
        print('\nERROR: No open workbooks to run the macro in.')
        return False

    def show_execution(self, flag=True):
        """Sets the visibility of the VBA macro exeuction.

        Keyword Arguments:
            flag (Boolean): Denotes visibility

        """
        if self.excel_app:
            self.excel_app.Visible = flag
            return True
        print('\nERROR: No open workbooks to set the visibility property.')
        return False
