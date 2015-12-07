Vycro
=====

`vycro` is a Python wrapper designed to make calling VBA functions and subroutines with Microsoft Excel as easy as a couple of standard file I/O calls.

``` python
from vycro import MacroWrapper

mw = MacroWrapper()
wb_kwargs = {"ReadOnly":"True"}
m_args = ["arg1", "True", "111"]
with mw.open_workbook("/full/path/with.extension", **wb_kwargs):
    mw.run_macro("macro_name", *m_args)  # See declaration for function args
```

`vycro` currently supports Python 2.7 and Python 3.1-3.5, and requires pywin32-219.
