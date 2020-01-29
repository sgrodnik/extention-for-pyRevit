# -*- coding: utf-8 -*-
""""""
__title__ = 'Выбрать\nприток'
__author__ = 'SG'

import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# t = Transaction(doc, 'Выбрать приток')
# t.Start()

k = 304.8

ducts = FilteredElementCollector(doc)\
.OfCategory(BuiltInCategory.OST_DuctCurves)\
.WhereElementIsNotElementType().ToElements()

flexDucts = FilteredElementCollector(doc)\
.OfCategory(BuiltInCategory.OST_FlexDuctCurves)\
.WhereElementIsNotElementType().ToElements()

ductFittings = FilteredElementCollector(doc)\
.OfCategory(BuiltInCategory.OST_DuctFitting)\
.WhereElementIsNotElementType().ToElements()

ductAccessories = FilteredElementCollector(doc)\
.OfCategory(BuiltInCategory.OST_DuctAccessory)\
.WhereElementIsNotElementType().ToElements()

lst = list(ducts) + list(flexDucts) + list(ductFittings) + list(ductAccessories)

arr = [663639,666083,663639,666090,663639,666095,663639,666116,663733,666083,663733,666095,663928,665853,663928,665939,663928,666083,663952,666083,663952,666095,664041,666095,664274,665853,664530,666090,664539,666090,664541,666090,664541,666178,664546,666090
]
uidoc.Selection.SetElementIds(List[ElementId]([ElementId(i) for i in arr]))

# result = []
# for el in lst:
#     if el.LookupParameter('Классификация систем').AsString() == 'Приточный воздух':
#         result.append(el.Id)

# uidoc.Selection.SetElementIds(List[ElementId](result))

# t.Commit()