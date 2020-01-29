# -*- coding: utf-8 -*-
"""Нумератор пространств"""
__title__ = 'Нумеровать\nпространство'
__author__ = 'SG'
import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId
import sys
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# tags = [doc.GetElement(id) for id in uidoc.Selection.GetElementIds() if isinstance(doc.GetElement(id), IndependentTag)]

# if not tags:
#     sys.exit()


# class CustomISelectionFilter(ISelectionFilter):
#     def __init__(self, nom_categorie):
#         self.nom_categorie = nom_categorie

#     def AllowElement(self, e):
#         if self.nom_categorie in e.Category.Name:
#             return True
#         else:
#             return False

#     def AllowReference(self, ref, point):
#         return true


# i = 1
# try:
#     target = uidoc.Selection.PickObject(ObjectType.Element, CustomISelectionFilter("Пространства"), 'Выберите пространство')
#     target = doc.GetElement(target.ElementId)
#     i = int(target.LookupParameter('Номер').AsString()) + 1
# except:  # Exceptions.OperationCanceledException:
#     pass

# t = Transaction(doc, 'Перенести марки')
# t.Start()

# while 1:
#     try:
#         target = uidoc.Selection.PickObject(ObjectType.Element, CustomISelectionFilter("Пространства"), 'Выберите пространство')
#     except:  # Exceptions.OperationCanceledException:
#         # t.Commit()
#         break
#         # sys.exit()

#     target = doc.GetElement(target.ElementId)
#     target.LookupParameter('Номер').Set(str(i))
#     i += 1

t = Transaction(doc, 'Нумератор')
t.Start()

sel = [doc.GetElement(elid) for elid in uidoc.Selection.GetElementIds()]

i = 1
for el in sel:
    el.Text = str(i)
    # el.LookupParameter('Текст').Set(str(i))
    i += 1


t.Commit()
