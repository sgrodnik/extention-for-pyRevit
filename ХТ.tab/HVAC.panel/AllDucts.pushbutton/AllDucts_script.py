# -*- coding: utf-8 -*-
""""""
__title__ = 'Выбрать\nвоздуховоды'
__author__ = 'SG'

import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

k = 304.8

ducts = FilteredElementCollector(doc)\
.OfCategory(BuiltInCategory.OST_DuctCurves)\
.WhereElementIsNotElementType().ToElements()


uidoc.Selection.SetElementIds(List[ElementId]([i.Id for i in ducts]))
