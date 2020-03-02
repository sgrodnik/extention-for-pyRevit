# -*- coding: utf-8 -*-
#pylint: disable=W0703,E0401,C0103,C0111
from pyrevit import revit, DB
from pyrevit import forms
from pyrevit import script
import re

__title__ = 'Перенумерация\nлистов'
__author__ = 'PyRevit, SG'

# __doc__ = 'Increases the sheet number of the selected sheets by one. '\
#           'The sheet name change will be printed if logging is set '\
#           'to Verbose in pyRevit settings.'

selection = [i for i in revit.get_selection() if i.LookupParameter('Категория').AsValueString() == 'Листы']

# selected_sheets = forms.select_sheets(title='Select Sheets') if not selection else selection
# selected_sheets or script.exit()

# # for i in selected_sheets:
# #     print(i.SheetNumber)
# # print(111)


# from pyrevit import forms

# q = forms.ask_for_string(
#     default='some-tag',
#     prompt='Enter new tag name:',
#     title='Tag Manager'
# )

# print(q)

def natural_sorted(list, key=lambda s: s):
    """
    Sort the list into natural alphanumeric order.
    """
    def get_alphanum_key_func(key):
        convert = lambda text: int(text) if text.isdigit() else text
        return lambda s: [convert(c) for c in re.split('([0-9]+)', key(s))]
    sort_key = get_alphanum_key_func(key)
    return sorted(list, key=sort_key)


sorted_sheet_list = natural_sorted(selected_sheets, key=lambda x: x.SheetNumber)
data = []
for sheet in sorted_sheet_list:
    da = []
    da.append(sheet.Id)
    da.append(sheet.Number)
    da.append(sheet.Name)
    da.append('\t'.join(sheet.Number.split('.')))
    data.append(da)


sorted_sheet_list.reverse()
# for i in sorted_sheet_list:
#     print(i.SheetNumber)
# print(111)

# with revit.TransactionGroup('Перенумерация листов') as t:
#     with revit.Transaction('Перенумерация листа'):
#         for sheet in sorted_sheet_list:
#                 try:
#                     cur_sheet_num = sheet.SheetNumber
#                     sheetnum_p = sheet.Parameter[DB.BuiltInParameter.SHEET_NUMBER]
#                     # print(str(int(sheet.SheetNumber) + 1))
#                     sheetnum_p.Set(str(int(sheet.SheetNumber) + 1))
#                     # new_sheet_num = sheet.SheetNumber
#                     # logger.info('{} -> {}'.format(cur_sheet_num, new_sheet_num))
#                 except Exception as shift_err:
#                     print(shift_err)
#                     t.RollBack()
#         revit.doc.Regenerate()
