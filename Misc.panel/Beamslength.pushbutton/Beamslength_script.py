# -*- coding: utf-8 -*-
""""""
__title__ = 'Длина\nбалок'
__author__ = 'SG'
import os
import shutil
import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *
from Autodesk.Revit.DB import Structure, FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId
import sys
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

k = 304.8

# els = [doc.GetElement(id) for id in uidoc.Selection.GetElementIds()]

els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).WhereElementIsNotElementType().ToElements()

t = Transaction(doc, 'Длина балок')
t.Start()

for el in els:
    el.LookupParameter('Длина факт').Set('{:.0f}'.format(el.LookupParameter('Фактическая длина').AsDouble() * k))
    dimB = el.LookupParameter('ADSK_Размер_Высота').AsDouble() * k
    dimH = el.LookupParameter('ADSK_Размер_Ширина').AsDouble() * k
    el.LookupParameter('Наименование').Set('{:.0f}×{:.0f}'.format(min([dimB, dimH]), max([dimB, dimH])))

    if not el.LookupParameter('Этап').HasValue:
    	el.LookupParameter('Этап').Set(999 * 10.76391041671)

t.Commit()
