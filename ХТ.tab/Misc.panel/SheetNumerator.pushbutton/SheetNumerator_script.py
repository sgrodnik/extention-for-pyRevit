# -*- coding: utf-8 -*-
"""Номера с точкой или с буквами: части после точки и буквы не будут изменены.
Номера с запятой: сменятся целими числами (части после запятой будут стёрты)"""
__title__ = 'Нумеровать\nтекст'
__author__ = 'SG'

import clr
import re
clr.AddReference('System.Core')
from System.Collections.Generic import *
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

sel = [el for el in [doc.GetElement(id) for id in uidoc.Selection.GetElementIds()] if el.Category.Name == 'Текстовые примечания']
num = 1
if __shiftclick__:
    from pyrevit import forms

    num = forms.ask_for_string(
        default='1',
        prompt='Введите начальное значение',
        title='Нумеровать текст'
    )
    num = int(num)
t = Transaction(doc, "Нумеровать текст")
t.Start()

for el in sel:
    el.Text = str(num)
    num += 1

t.Commit()

# def natural_sorted(list, key=lambda s: s):
#     """
#     Sort the list into natural alphanumeric order.
#     """
#     def get_alphanum_key_func(key):
#         def convert(text): return int(text) if text.isdigit() else text
#         return lambda s: [convert(c) for c in re.split('([0-9]+)', key(s))]
#     sort_key = get_alphanum_key_func(key)
#     return sorted(list, key=sort_key)


# sheets = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets).ToElements()
# sheets = [i for i in sheets
#           if i.LookupParameter('Имя листа').AsString() != 'Начальный вид']

# sel = [el for el in [doc.GetElement(id) for id in uidoc.Selection.GetElementIds()] if el.Category.Name == 'Листы']

# if len(sel) > 1:
#     sheets = sel

# # for i in natural_sorted(sheets, key = lambda x: x.SheetNumber.replace('‎', '')):
# #   print(i.SheetNumber)

# ###########################################

# digits = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']


# def splitNumber(number):
#     number = number.replace('.', 'dot').replace(',', '.')
#     i = 0
#     for char in number:
#         if char == '‎':
#             i += 1
#         else:
#             break
#     nameWithoutPrefix = number[i:]
#     j = 0
#     for char in nameWithoutPrefix:
#         if char in digits or char == '.':
#             j += 1
#         else:
#             break
#     return number[:i], nameWithoutPrefix[:j], nameWithoutPrefix[j:].replace('dot', '.')


# class sheet():
#     def __init__(self, revitsSheet):
#         self.reference = revitsSheet
#         self.oldNumber = revitsSheet.SheetNumber.replace('temp', '')
#         self.prefix, self.number, self.suffix = splitNumber(revitsSheet.SheetNumber.replace('temp', ''))

#     def __str__(self):
#         return 'mySheet {}/{}/{} : {}/{}/{}'.format(self.prefix, self.number, self.suffix, len(self.prefix), len(self.number), len(self.suffix))

# parts = {}
# for s in sheets:
#     if s.LookupParameter('ADSK_Штамп Раздел проекта'):
#         part = s.LookupParameter('ADSK_Штамп Раздел проекта').AsString()
#     else:
#         part = s.LookupParameter('Раздел проекта').AsString()
#     if part not in parts:
#         parts[part] = []
#     parts[part].append(s)

# t = Transaction(doc, "Нумеровать листы")
# t.Start()

# j = 1
# for part in sorted(parts.keys()):
#     # print('\n-------------- ' + part + ' --------------')

#     mySheets = []
#     for i in natural_sorted(parts[part], key=lambda x: x.SheetNumber.replace('‎', '')):
#         sh = sheet(i)
#         mySheets.append(sh)

#     for sh in mySheets:
#         j = j + 1
#         sh.reference.SheetNumber += 'temp'

# for part in sorted(parts.keys()):
#     print('\n-------------- ' + (part if part else '???') + ' --------------')

#     mySheets = []
#     for i in natural_sorted(parts[part], key=lambda x: x.SheetNumber.replace('‎', '')):
#         sh = sheet(i)
#         mySheets.append(sh)
#     # print(111)
#     # print(mySheets[0].number)
#     # print(222)
#     firstRun = 1
#     i = float(mySheets[0].number)
#     for sh in mySheets:
#         delta = 0 if sh.suffix else 1
#         if firstRun:
#             firstRun = 0
#             delta = 0
#         i = i + delta
#         new = mySheets[0].prefix + '{:.0f}'.format(i) + sh.suffix
#         old = sh.oldNumber.replace('‎', '·')
#         change = ' !' if new.replace('‎', '·') != old else ''
#         print('{}-> {}{}'.format(old.ljust(5, fillchar=' '), new.replace('‎', '·').ljust(5, fillchar=' '), change))
#         sh.reference.SheetNumber = new

# t.Commit()

# # for i in natural_sorted([
# #   '1',
# #   '1,1',
# #   '1,05',
# #   '1,2',
# #   '2',
# #   '2,1',
# #   '2,05',
# #   '3',
# #   '4',
# #   '5'
# #   ]):
# #   print(i)
# #########################################

# # lst = []
# # for i in sheets:
# #   sublist = ['' for s in '012345']
# #   sublist[0] = i
# #   sublist[1] = i.LookupParameter('Номер листа').AsString()
# #   if '-' in sublist[1]:
# #       sublist[1] = sublist[1][:sublist[1].find('-')]
# #   for index, pos in enumerate(sublist[1]):
# #       if pos.isdigit():
# #           digit = index
# #           break
# #   sublist[2] = sublist[1][:digit] # Текстовая часть
# #   sublist[3] = sublist[1][digit:] # Непосредственно номер
# #   lst.append(sublist)

# # def myint(s):
# #   if s.find(',') == -1 and s.find('.') == -1:
# #       return int(s)
# #   else:
# #       if s.find(',') >= 0:
# #           return int(s[:s.find(',')])
# #       else:
# #           return int(s[:s.find('.')])

# # lst.sort(key = lambda x: myint(x[3]))
# # lst.sort(key = lambda x: x[2])

# # for index, pos in enumerate(lst):
# #   if index == 0: # Первый раз просто берём номер
# #       pos[4] = pos[3]
# #   else:
# #       prev = lst[index-1][4]
# #       if '.' in prev: # Если предыдущий с индексом
# #           prev = int(prev[:prev.find('.')]) # Берём часть до индекса
# #           if '.' in pos[3]: # Если и в предыдущем и в текущем есть индекс
# #               current = prev # То не прибавляем единицу
# #           else:
# #               current = prev + 1 # Иначе - прибавляем
# #       else: # Если же предыдущий без индекса
# #           try:
# #               prev = int(prev)
# #           except:
# #               break
# #           current = prev + 1 # То прибавляем единицу
# #       if pos[2] != lst[index-1][2]:
# #           current = pos[3][:pos[3].find('.')] if pos[3][:pos[3].find('.')] else pos[3]
# #       if '.' in pos[3]:
# #           pos[4] = '{}.{}'.format(current, pos[3][pos[3].find('.')+1:])
# #       else:
# #           pos[4] = '{}'.format(current)

# # # els = []
# # # for i in lst:
# # #     els = FilteredElementCollector(doc, i[0].Id)
# # #     for e in els:
# # #         if doc.GetElement(e.GetTypeId()):
# # #             if 'тамп' in doc.GetElement(e.GetTypeId()).get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString():
# # #                 i[5] = '-' + doc.GetElement(e.GetTypeId()).get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()

# # els = []
# # for i in lst:
# #   els = FilteredElementCollector(doc, i[0].Id)
# #   for e in els:
# #       if 1:
# #           if 1:
# #               i[5] = ''

# # string = ' ----------- OK -----------\n'
# # for i in lst:
# #   for j in i:
# #       string += '{}'.format(j if type(j) == str else j.Name) + ", "
# #   string += '\n'

# # t = Transaction(doc, "Нумератор листов")
# # err = ''
# # t.Start()
# # for index, i in enumerate(lst):
# #   i[0].LookupParameter('Номер листа').Set('w' + i[1])
# # for i in lst:
# #   i[0].LookupParameter('Номер листа').Set(i[2] + i[4] + i[5])
# #   try: i[0].LookupParameter('ХТ Номер листа').Set(i[4])
# #   except: err = 'Параметр "ХТ Номер листа" не существует, значение не присвоено.\n'
# # t.Commit()

# # print(err + string)
