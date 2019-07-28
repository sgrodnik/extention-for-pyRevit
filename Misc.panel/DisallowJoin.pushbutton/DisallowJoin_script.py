# -*- coding: utf-8 -*-
""""""
__title__ = 'Отмена\nсоединений'
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

els = [doc.GetElement(id) for id in uidoc.Selection.GetElementIds()]

els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).WhereElementIsNotElementType().ToElements()

t = Transaction(doc, 'Отмена соединений')
t.Start()

for el in els:
    Structure.StructuralFramingUtils.DisallowJoinAtEnd(el, 0)
    Structure.StructuralFramingUtils.DisallowJoinAtEnd(el, 1)

t.Commit()
