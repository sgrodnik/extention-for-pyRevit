# -*- coding: utf-8 -*-
""""""
__title__ = 'Test'
__author__ = 'SG'

import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *
from Autodesk.Revit.DB import ElementId, PartUtils, ViewOrientation3D, XYZ, FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, Line, Structure
import sys
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
# from Autodesk.Revit.ApplicationServices.Application import Create
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application
k = 304.8

t = Transaction(doc, 'Test')

t.Start()

src = '''ОК-8   2100    900     ГОСТ 30674-99   ОП Б2 2100-900 (4М-12-4М-20-И4) 0
ОК-40   2100    2700        ГОСТ 30674-99   ОП Б2 2100-2700 (4М-12-4М-20-И4)    Глухое
ОК-37   2100    2700        ГОСТ 30674-99   ОП Б2 2100-2700 (4М-12-4М-12-И4)    0
ОК-31   2100    1800        ГОСТ 30674-99   ОП Б2 2100-1800 (4М-12-4М-20-И4)    Глухое
ОК-30   2100    1800        ГОСТ 30674-99   ОП Б2 2100-1800 (4М-12-4М-20-И4)    Глухое
ОК-26   2100    1800        ГОСТ 30674-99   ОП Б2 2100-1800 (4М-12-4М-12-И4)    0
ОК-14   2100    900     ГОСТ 30674-99   ОП Б2 2100-900 (4М-12-4М-12-И4) 0
ОК-10   2100    900     ГОСТ 30674-99   ОП Б2 2100-900  (4М-12-4М-20-И4)    Глухое
Ов-9    1310    1100        ТУ производителя    1310х1100 (h) Rg    внутреннее окно, однокамерный стеклопакет, рентгензащитное Pb=1мм, поставка Philips
ОК-23   2100    1500        ГОСТ 30674-99   ОП Б2 2100-1500 (4М-12-4М-20-И4)    Глухое
ОК-11   2100    900     ГОСТ 30674-99   ОП Б2 2100-900  (4М-12-4М-20-И4)    Глухое'''

raws = src.split('\n')

# els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Doors).WhereElementIsElementType().ToElements()

# sel = [doc.GetElement(elid) for elid in uidoc.Selection.GetElementIds()]

# sel_symbol = doc.GetElement(sel[0].GetTypeId())

# for raw in raws:
#   raw = raw.split('\t')
#   # if float(raw[2]) >= 1400:
#   if float(raw[2]) < 1400:
#       symbols = list(filter(lambda x: x.LookupParameter('Имя типа').AsString() == raw[0], els))
#       if symbols:
#           symbol = symbols[0]
#       else:
#           symbol = sel_symbol
#           symbol = symbol.Duplicate(raw[0])
#       symbol.LookupParameter('Наличие окна').Set(1500 / k if raw[1] == 'окно' else 15000 / k)
#       symbol.LookupParameter('Ширина').Set(float(raw[2]) / k)
#       symbol.LookupParameter('Комментарии к типоразмеру').Set(raw[6])
#       symbol.LookupParameter('ADSK_Марка').Set(raw[0])
#       symbol.LookupParameter('ADSK_Наименование').Set(raw[5])
#       symbol.LookupParameter('ADSK_Обозначение').Set(raw[4])
#       if symbol.LookupParameter('Левая'):
#           symbol.LookupParameter('Левая').Set(True if raw[3] == 'Л' else False)





# els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Windows).WhereElementIsElementType().ToElements()

# sel = [doc.GetElement(elid) for elid in uidoc.Selection.GetElementIds()]

# sel_symbol = doc.GetElement(sel[0].GetTypeId())

# for raw in raws:
#   raw = raw.split('\t')
#   symbols = list(filter(lambda x: x.LookupParameter('Имя типа').AsString() == raw[0], els))
#   if symbols:
#       symbol = symbols[0]
#   else:
#       symbol = sel_symbol
#       symbol = symbol.Duplicate(raw[0])
#   symbol.LookupParameter('Примерная ширина').Set((float(raw[2]) - 40) / k)
#   symbol.LookupParameter('Примерная высота').Set((float(raw[1]) - 50) / k)
#   symbol.LookupParameter('Комментарии к типоразмеру').Set(raw[6] if raw[6] != '0' else '')
#   symbol.LookupParameter('ADSK_Марка').Set(raw[0])
#   symbol.LookupParameter('ADSK_Наименование').Set(raw[5])
#   symbol.LookupParameter('ADSK_Обозначение').Set(raw[4])

















# src = '''4001 Шлюз для пациентов при входе в операционный блок, с перекладчиком   17,18   3
# 4002  Коридор 109,7   18
# 4022  Наркозная   13,03   2
# 4023  Предоперационная    12,02   2
# 4024  Наркозная   12,39   2
# 4025  Операционная общей хирургии, урологии, травматологии и ортопедии (БП)   45,36   8
# 4027  Операционная ЛОР и ЧЛХ (БП) 36,02   6
# 4042  Шлюз эвак.  10,21   2
# 4050  Палата преднаркозная/пробуждения на 3 койки с постом медсестры  45,28   8
# 4061  Шлюз    6,27    1
# 4062  Коридор 99,24   17
# 4069  Шлюз    5,95    1
# 4083  Палата реанимации и интенсивной терапии на 6 коек с постом медсестры (в том числе 1 койка кардиологии) - (БП)   87,4    15
# 7001  Шлюз    7,91    1
# 7002  Коридор 129,49  22
# 7003  Наркозная   12,15   2
# 7004  Предоперационная    15,83   3
# 7005  Наркозная   12,6    2
# 7006  Операционная для общей хирургии, интегрированная - (БП) 42,34   7
# 7007  Операционная общепрофильная для гинекологии и урологии, интегрированная - (БП)  36,19   6
# 7008  Коридор 96,83   16
# 7010  Техническое помещение операционной  11,21   2
# 7011  Предоперационная    13,22   2
# 7012  Наркозная   12,05   2
# 7013  Комната управления  30,33   5
# 7014  Наркозная   17,11   3
# 7015  Операционная гибридная для сердечно-сосудистой хирургии - (БП)  66,98   11
# 7017  Операционная гибридная для нейрохирургии - (БП) 68,54   11
# 7018  Предоперационная    21,28   4
# 7019  Техническое помещение операционной  11,37   2
# 7021  Коридор 111,18  19
# 7036  Шлюз    10,94   2
# 7049  Наркозная   12,03   2
# 7050  Предоперационная    12,06   2
# 7051  Операционная общепрофильная для ЛОР и проктологии - (БП)    42,9    7
# 7053  Наркозная   12,4    2
# 7054  Операционная ортопедо-травматологическая - (БП) 43,14   7
# 7057  Палата пробуждения на 1 койку с постом медсестры    22,69   4
# 7083  Коридор 23,66   4
# 7091  Шлюз пациента   14,7    2
# 7100  Коридор 27,76   5
# 7103  Пост медсестер и врачей с мониторами    15,99   3
# 7104  Палата реанимации и интенсивной терапии на 8 коек - (БП) с постом медсестры 186,32  31
# 7110  Коридор 82,1    14
# 7133  Палата пробуждения на 5 коек с постом медсестры 87,6    15
# 7140  Коридор 32,91   5'''

# raws = src.split('\n')


# els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()
# for el in els:
#   for raw in raws:
#       raw = raw.split('\t')
#       if raw[0] == el.LookupParameter('Номер').AsString():
#           number = raw[3]
#   el.LookupParameter('Комментарии').Set(number)








t.Commit()
