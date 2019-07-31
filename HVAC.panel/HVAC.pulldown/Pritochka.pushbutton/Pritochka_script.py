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

t = Transaction(doc, 'Выбрать приток')
t.Start()

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

result = []
for el in lst:
	if el.LookupParameter('Классификация систем').AsString() == 'Приточный воздух':
		result.append(el.Id)

uidoc.Selection.SetElementIds(List[ElementId](result))

t.Commit()