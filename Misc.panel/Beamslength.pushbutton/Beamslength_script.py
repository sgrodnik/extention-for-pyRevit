# -*- coding: utf-8 -*-
""""""
__title__ = 'Выбрать\nэтап'
__author__ = 'SG'
import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId, XYZ, Structure

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).WhereElementIsNotElementType().ToElements()
sel = [doc.GetElement(elid) for elid in uidoc.Selection.GetElementIds()]
sel = [i.LookupParameter('Этап').AsDouble() for i in sel if i.LookupParameter('Категория').AsValueString() == 'Каркас несущий']
el_ids = List[ElementId]([i.Id for i in els if i.LookupParameter('Этап').AsDouble() in sel])

uidoc.Selection.SetElementIds(el_ids)