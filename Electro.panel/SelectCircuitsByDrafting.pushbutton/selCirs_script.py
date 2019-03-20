# -*- coding: utf-8 -*-
"""Выбрать элементы цепей (розетки, коробки, щиты), соответствующих выбранным УГО на однолинейной схеме"""
__title__ = 'Выбрать\nэлементы'
__author__ = 'test'

import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

sel = uidoc.Selection.GetElementIds()
if len(sel) != 0:
    arr = []
    for el in [doc.GetElement(id) for id in sel]:
        for i in el.LookupParameter('Цепи').AsString().split():
            id = ElementId(int(i))
            cir = doc.GetElement(id)
            arr.extend(list(cir.Elements))
            arr.append(cir.BaseEquipment)
    arr = list(set([i.Id for i in arr]))

    uidoc.Selection.SetElementIds(List[ElementId](arr))

# Дописать щиты