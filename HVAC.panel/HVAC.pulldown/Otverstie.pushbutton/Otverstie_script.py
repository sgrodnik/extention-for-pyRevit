# -*- coding: utf-8 -*-
"""Otverstie"""
__title__ = 'Otverstie'
__author__ = 'SG'

import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

t = Transaction(doc, 'Otverstie')
t.Start()

k = 304.8

ducts = FilteredElementCollector(doc)\
.OfCategory(BuiltInCategory.OST_DuctCurves)\
.WhereElementIsNotElementType().ToElements()

for duct in ducts:
	name = duct.LookupParameter('Имя системы').AsString()
	if duct.LookupParameter('Ширина'):
		a = duct.LookupParameter('Ширина').AsDouble() * k + 60
		b = duct.LookupParameter('Высота').AsDouble() * k + 60
		duct.LookupParameter('Отверстие').Set('{}: {:.0f}x{:.0f}'.format(name, a, b))

t.Commit()