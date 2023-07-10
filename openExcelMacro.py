import os
from time import time
import time
from win32com import client
EXCEL_CLS_NAME = "Excel.Application"
class XlMacro:
    def __init__(self, path, book, module, name, *args):
        self._path = path  # path containing workbook
        self._book = book  # workbook name like Book1.xlsm
        self._module =  module  # module name, e.g., Module1
        self._name = name  # procedure or function name
        self._params = args  # argument(s)
        self._wb = None
    @property
    def workbook(self):
        return self._wb
    @property
    def wb_path(self):
        return os.path.join(self._path, self._book)
    @property
    def name(self):
        return f'{self._book}!{self._module}.{self._name}'
    @property
    def params(self):
        return self._params
    def get_workbook(self):
        wb_name = os.path.basename(self.wb_path)
        try:
            xl = client.GetActiveObject(EXCEL_CLS_NAME)
        except:
            # Excel is not running, so we need to handle it.
            xl = client.Dispatch(EXCEL_CLS_NAME)
        if wb_name in [wb.Name for wb in xl.Workbooks]:
            return xl.Workbooks[wb_name]
        else:
            return xl.Workbooks.Open(self.wb_path)
    def Run(self, *args, **kwargs):
        """ 
        Runs an Excel Macro or evaluates a UDF 
        returns None for macro, or the return value of Function
        NB: there is no error-handling here, but there should be!
        """
        keep_open = kwargs.get('keep_open', True)
        save_changes = kwargs.get('save_changes', False)
        self._wb = self.get_workbook()
        xl_app = self._wb.Application
        xl_app.Visible = True
        ret = None
        if args is None:
            ret = xl_app.Run(self.name)
        elif not args:
            # run with some default parameters
            ret = xl_app.Run(self.name, *self.params)
        else:
            ret = xl_app.Run(self.name, *args)
        if not keep_open:
            self.workbook.Close(save_changes)
            self._wb = None
            xl_app.Quit()
        return ret

# Modify these path/etc as needed.
path = r'C:\EXCEL Altxlstart' 
book = 'PERSONAL.XLSB'
module = "Python"
macros = ['pythonOpenFromServer', 'pythonImportRowsFromWorksheetOwsEstados', 'pythonImportRowsFromWorksheetOwsControlTickets',] # pythonImportRowsFromOutboundToBulkSheetAndNoStock  
def default_params(macro):
    """
    mocks some dummy arguments for each Excel macro
    this is required by Excel.Application.Run(<method>,<ParamArray>)
    """
    d = {'macro1': ("hello", "world", 123, 18.45),
        'SayThis': ('hello, world!',),
        'AddThings': [13, 39],
        'GetFromExcel': [],
        'GetWithArguments': [2]
        }
    return d.get(macro)
# Run the macros and their arguments:
for m in macros:
    args = default_params(m)
    if args:
        macro = XlMacro(path, book, module, m, *args)
    else:
        macro = XlMacro(path, book, module, m)
    x = macro.Run()
    time.sleep(5)
    print(f'returned {x} from {m}({args})' if x else f'Successfully executed {m}({args})')


