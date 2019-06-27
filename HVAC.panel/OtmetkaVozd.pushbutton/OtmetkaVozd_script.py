# -*- coding: utf-8 -*-
"""Отметка воздуховода"""
__title__ = 'Отметка\nвоздуховода'
__author__ = 'SG'

import clr
clr.AddReference('System.Core')
import sys
from System.Collections.Generic import *

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


class CustomISelectionFilter(ISelectionFilter):
    def __init__(self, nom_categorie):
        self.nom_categorie = nom_categorie

    def AllowElement(self, e):
        if self.nom_categorie in e.Category.Name:
            return True
        else:
            return False

    def AllowReference(self, ref, point):
        return true


sel = [doc.GetElement(id) for id in uidoc.Selection.GetElementIds()]

if not sel:
    sys.exit()

try:
    target = uidoc.Selection.PickObject(ObjectType.Element, CustomISelectionFilter('Воздуховоды'), 'Укажите воздуховод с нужной отметкой')
    target = doc.GetElement(target.ElementId)
except:  # Exceptions.OperationCanceledException:
    sys.exit()

t = Transaction(doc, 'Отметка воздуховода')
t.Start()

for el in sel:
    el.LookupParameter('Смещение').Set(target.LookupParameter('Смещение').AsDouble())

#k = 304.8

t.Commit()
