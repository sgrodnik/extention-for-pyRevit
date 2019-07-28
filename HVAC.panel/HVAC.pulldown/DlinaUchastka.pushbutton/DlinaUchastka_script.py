# -*- coding: utf-8 -*-
"""Otverstie"""
__title__ = 'Длины\nучастков'
__author__ = 'SG'

import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

t = Transaction(doc, 'Длины участков')
t.Start()

k = 304.8

pipeCurves = FilteredElementCollector(doc)\
    .OfCategory(BuiltInCategory.OST_PipeCurves)\
    .WhereElementIsNotElementType().ToElements()

pipeFittings = FilteredElementCollector(doc)\
    .OfCategory(BuiltInCategory.OST_PipeFitting)\
    .WhereElementIsNotElementType().ToElements()

pipeAccessories = FilteredElementCollector(doc)\
    .OfCategory(BuiltInCategory.OST_PipeAccessory)\
    .WhereElementIsNotElementType().ToElements()

els = []
for i in pipeCurves:
    els.append(i)
for i in pipeFittings:
    els.append(i)
for i in pipeAccessories:
    els.append(i)

dict = {}
for el in els:
    len = el.LookupParameter('Длина').AsDouble() * k if el.LookupParameter('Длина') else 0
    sector = el.LookupParameter('Участок').AsString()
    if sector not in dict:
        dict[sector] = 0
    dict[sector] += len

for el in els:
    sector = el.LookupParameter('Участок').AsString()
    len = dict[sector]
    el.LookupParameter('Длина участка').Set(len/1000)

# for duct in ducts:
#     name = duct.LookupParameter('Имя системы').AsString()
#     if duct.LookupParameter('Ширина'):
#         a = duct.LookupParameter('Ширина').AsDouble() * k + 60
#         b = duct.LookupParameter('Высота').AsDouble() * k + 60
#         duct.LookupParameter('Отверстие').Set('{}: {:.0f}x{:.0f}'.format(name, a, b))

t.Commit()
