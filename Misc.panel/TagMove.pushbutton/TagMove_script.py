# -*- coding: utf-8 -*-
"""Выравнивание марок"""
__title__ = 'Перенести\nмарки'
__author__ = 'SG'
import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *
from Autodesk.Revit.DB import IndependentTag, XYZ, FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId
import sys
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

tags = [doc.GetElement(id) for id in uidoc.Selection.GetElementIds() if isinstance(doc.GetElement(id), IndependentTag)]

if not tags:
    sys.exit()

class CustomISelectionFilter(ISelectionFilter):
    def __init__(self, nom_categorie):
        self.nom_categorie = nom_categorie

    def AllowElement(self, e):
        if e.Category.Name == self.nom_categorie:
            return True
        else:
            return False

    def AllowReference(self, ref, point):
        return true

try:
    target = uidoc.Selection.PickObject(ObjectType.Element, CustomISelectionFilter("Марки стен"), 'Выберите целевую марку')
except:  # Exceptions.OperationCanceledException:
    sys.exit()

target = doc.GetElement(target.ElementId)

t = Transaction(doc, 'Перенести марки')
t.Start()

if target.LeaderElbow.X < target.TagHeadPosition.X:
    for tag in tags:
        if tag.LeaderElbow.X < tag.TagHeadPosition.X:
            tag.LeaderElbow = target.LeaderElbow
            tag.TagHeadPosition = target.TagHeadPosition
        else:
            tag.LeaderElbow -= tag.TagHeadPosition - target.TagHeadPosition
            tag.TagHeadPosition = target.TagHeadPosition
else:
    for tag in tags:
        if tag.LeaderElbow.X > tag.TagHeadPosition.X:
            tag.LeaderElbow = target.LeaderElbow
            tag.TagHeadPosition = target.TagHeadPosition
        else:
            tag.LeaderElbow -= tag.TagHeadPosition - target.TagHeadPosition
            tag.TagHeadPosition = target.TagHeadPosition

t.Commit()
