# -*- coding: utf-8 -*-
""""""
__title__ = 'Test'
__author__ = 'SG'

import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *
from Autodesk.Revit.DB import SectionType, ElementId, PartUtils, ViewOrientation3D, XYZ, FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, Line, Structure
import sys
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
# from Autodesk.Revit.ApplicationServices.Application import Create
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application
k = 304.8



t = Transaction(doc, 'Test')

t.Start()

cirs = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalCircuit).WhereElementIsNotElementType().ToElements()


sel = [doc.GetElement(elid) for elid in uidoc.Selection.GetElementIds()][0]
print(sel.PipeSegment)



t.Commit()








# schedules = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Schedules).WhereElementIsNotElementType().ToElements()

# schedules = list(filter(lambda x: 'Пом ' in x.Name, schedules))

# # print(schedules)

# for sch in sorted(schedules, key=lambda x: x.Name):
#   print(sch.Name)
#   for sch_filter in sch.Definition.GetFilters():
#       print(sch_filter.FieldId)
#       print(sch_filter.FilterType)
#       if sch_filter.IsDoubleValue:
#           print(sch_filter.GetDoubleValue())
#       elif sch_filter.IsElementIdValue:
#           print(sch_filter.GetElementIdValue())
#       elif sch_filter.IsIntegerValue:
#           print(sch_filter.GetIntegerValue())
#       elif sch_filter.IsStringValue:
#           print(sch_filter.GetStringValue())
#       elif sch_filter.IsNullValue:
#           print('Null')
#       else:
#           print('--------------else')
#   print()

    # sch_filter = sch.Definition.GetFilters()[0]
    # sch_filter.SetValue(sch.Name)
    # sch.Definition.SetFilter(0, sch_filter)


# print(schedule.Definition.GetFilters()[0].GetStringValue())
# filter = schedule.Definition.GetFilters()[0]
# filter.SetValue('111')
# schedule.Definition.SetFilter(0,filter)
# print(schedule.Definition.GetFilters()[0].GetStringValue())

# schedule_filter = schedule.Definition.GetFilters()[0]
# print(schedule_filter.FieldId)






# # sel = [doc.GetElement(elid) for elid in uidoc.Selection.GetElementIds()]

# # els = FilteredElementCollector(doc, ElementId(2446)).ToElements()

# # schedule = sel[0].GetTableData().GetSectionData(SectionType.Body)
# # schedule.InsertRow(schedule.FirstRowNumber + 1)
# # schedule.SetCellText(0, 0, '123')

# # for el in els:
# #     el.LookupParameter('й1').Set('1q')

# # print(els)


# # data = schedule.GetTableData().GetSectionData(SectionType.Body)

# # # for i in range(data.LastRowNumber - 0):
# # #   data.RemoveRow(i+1)

# # data.InsertRow(0)

# # # print(data.LastRowNumber)





# src = '''ОК-8   2100    900     ГОСТ 30674-99   ОП Б2 2100-900 (4М-12-4М-20-И4) 0
# ОК-40   2100    2700        ГОСТ 30674-99   ОП Б2 2100-2700 (4М-12-4М-20-И4)    Глухое
# ОК-37   2100    2700        ГОСТ 30674-99   ОП Б2 2100-2700 (4М-12-4М-12-И4)    0
# ОК-31   2100    1800        ГОСТ 30674-99   ОП Б2 2100-1800 (4М-12-4М-20-И4)    Глухое
# ОК-30   2100    1800        ГОСТ 30674-99   ОП Б2 2100-1800 (4М-12-4М-20-И4)    Глухое
# ОК-26   2100    1800        ГОСТ 30674-99   ОП Б2 2100-1800 (4М-12-4М-12-И4)    0
# ОК-14   2100    900     ГОСТ 30674-99   ОП Б2 2100-900 (4М-12-4М-12-И4) 0
# ОК-10   2100    900     ГОСТ 30674-99   ОП Б2 2100-900  (4М-12-4М-20-И4)    Глухое
# Ов-9    1310    1100        ТУ производителя    1310х1100 (h) Rg    внутреннее окно, однокамерный стеклопакет, рентгензащитное Pb=1мм, поставка Philips
# ОК-23   2100    1500        ГОСТ 30674-99   ОП Б2 2100-1500 (4М-12-4М-20-И4)    Глухое
# ОК-11   2100    900     ГОСТ 30674-99   ОП Б2 2100-900  (4М-12-4М-20-И4)    Глухое'''

# raws = src.split('\n')

# # els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Doors).WhereElementIsElementType().ToElements()

# # sel = [doc.GetElement(elid) for elid in uidoc.Selection.GetElementIds()]

# # sel_symbol = doc.GetElement(sel[0].GetTypeId())

# # for raw in raws:
# #   raw = raw.split('\t')
# #   # if float(raw[2]) >= 1400:
# #   if float(raw[2]) < 1400:
# #       symbols = list(filter(lambda x: x.LookupParameter('Имя типа').AsString() == raw[0], els))
# #       if symbols:
# #           symbol = symbols[0]
# #       else:
# #           symbol = sel_symbol
# #           symbol = symbol.Duplicate(raw[0])
# #       symbol.LookupParameter('Наличие окна').Set(1500 / k if raw[1] == 'окно' else 15000 / k)
# #       symbol.LookupParameter('Ширина').Set(float(raw[2]) / k)
# #       symbol.LookupParameter('Комментарии к типоразмеру').Set(raw[6])
# #       symbol.LookupParameter('ADSK_Марка').Set(raw[0])
# #       symbol.LookupParameter('ADSK_Наименование').Set(raw[5])
# #       symbol.LookupParameter('ADSK_Обозначение').Set(raw[4])
# #       if symbol.LookupParameter('Левая'):
# #           symbol.LookupParameter('Левая').Set(True if raw[3] == 'Л' else False)





# # els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Windows).WhereElementIsElementType().ToElements()

# # sel = [doc.GetElement(elid) for elid in uidoc.Selection.GetElementIds()]

# # sel_symbol = doc.GetElement(sel[0].GetTypeId())

# # for raw in raws:
# #   raw = raw.split('\t')
# #   symbols = list(filter(lambda x: x.LookupParameter('Имя типа').AsString() == raw[0], els))
# #   if symbols:
# #       symbol = symbols[0]
# #   else:
# #       symbol = sel_symbol
# #       symbol = symbol.Duplicate(raw[0])
# #   symbol.LookupParameter('Примерная ширина').Set((float(raw[2]) - 40) / k)
# #   symbol.LookupParameter('Примерная высота').Set((float(raw[1]) - 50) / k)
# #   symbol.LookupParameter('Комментарии к типоразмеру').Set(raw[6] if raw[6] != '0' else '')
# #   symbol.LookupParameter('ADSK_Марка').Set(raw[0])
# #   symbol.LookupParameter('ADSK_Наименование').Set(raw[5])
# #   symbol.LookupParameter('ADSK_Обозначение').Set(raw[4])

















# src = '''0.701    Лаборатория биоинформатики  27,75       5   29,14
# 0.702 Переговорная    29,53       5   31,01
# 0.703 Комната персонала   20,88       5   21,92
# 0.704 Кабинет заведующего 13,87       5   14,56
# 0.704a    Тамбур  3,2     5   3,36
# 0.705 Зона приема биоматериалов   19,68       5   20,66
# 0.706 Лаборатория ПГД 31,9        5   33,50
# 0.706a    Шлюз    5,55        5   5,83
# 0.707 Цитогенетическая лаборатория    34,97       5   36,72
# 0.707a    Шлюз    5,65        5   5,93
# 0.708 Лаборатория клеточных культур   25,84   1   15  29,72
# 0.708a    Шлюз    4,58        5   4,81
# 0.709 Лаборатория клеточной микроскопии   43,48       5   45,65
# 0.710 Лаборатория общего секвенирования   32,25   1   15  37,09
# 0.711 Лаборатория Pre-PCR 20,68   1   15  23,78
# 0.712 Лаборатория RT-PCR  14,15   1   15  16,27
# 0.713 Лаборатория PCR 9,35    1   15  10,75
# 0.714 Молекулярно-генетическая лаборатория    41,26   1   15  47,45
# 0.714a    Шлюз    3,64        5   3,82
# 0.715 Лаборатория электрофоретических исследований    19,31   1   15  22,21
# 0.715a    Шлюз    4,54        5   4,77
# 0.716 Лаборатория микрофлюидики   23,87       5   25,06
# 0.717 Зона подготовки проб    43,28       5   45,44
# 0.718 Лаборатория масс-спектрометрии  46,65       5   48,98
# 0.718а    Техническое помещение   3,13        5   3,29
# 0.720 Материальная    24,37       5   25,59
# 0.721 Помещение для временного хранения отходов   6       5   6,30
# 0.722 Помещение для приготовления растворов, чистой воды  15  1   15  17,25
# 0.723 Стерилизационная    12,26   1   15  14,10
# 0.724 Криохранилище   12,6        5   13,23
# 0.725 Помещение уборочного инвентаря  4       5   4,20
# 0.726.1   Санпропускник 1 7,46        5   7,83
# 0.726.2   Душевая 6,96        5   7,31
# 0.726.3   Санпропускник 2 7,15        5   7,51
# 0.727 Туалет  2,11        5   2,22
# 0.728 Коридор 173,71      5   182,40
# 0.792 Туалет  2,03        5   2,13
# 0.793 Туалет  2,16        5   2,27
# 0.029 Помещение криохранилища (-80 градусов Цельсия)  45,21   1   15  51,99
# 0.030 Помещение криохранилища (-196 градусов Цельсия) 26,86       5   28,20
# 0.031 Лаборатория аликвотирования 39,75   1   15  45,71
# 0.031a    Шлюз    4,02        5   4,22
# 0.032 Автоклавная 4,79        5   5,03
# 0.033 Техническое помещение   2,42        5   2,54
# 0.034 Упаковка чистых инструментов    7,25        5   7,61
# 0.035 Лаборатория геномного секвенирования    45,27       5   47,53
# 0.036 Кабинет научных сотрудников 25,74       5   27,03
# 0.037 Зона для переодевания персонала с санузлом  32,76       5   34,40
# 0.038 Помещение хранения специальной одежды   9,61        5   10,09
# 0.039 Материальная    15      5   15,75
# 0.040 Временное хранение отходов  6,02        5   6,32
# 0.041 Помещение уборочного инвентаря  5,8     5   6,09
# 0.042 Коридор 91,76       5   96,35'''

# rows = src.split('\n')


# els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()
# for index, el in enumerate(els):
#   row = rows[index].split('\t')
#   # if row[0] == el.LookupParameter('Номер').AsString():
#   #     number = row[3]
#   el.LookupParameter('Номер').Set(row[0])
#   el.LookupParameter('Имя').Set(row[1])
#   el.LookupParameter('Старая площадь число').Set(float(row[2].replace(',', '.')))
#   el.LookupParameter('Целевая площадь число').Set(float(row[5].replace(',', '.')))



# t.Commit()

