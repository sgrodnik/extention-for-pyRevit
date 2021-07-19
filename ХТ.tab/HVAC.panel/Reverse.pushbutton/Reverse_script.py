# -*- coding: utf-8 -*-
""""""
__title__ = 'Обернуть\nразмеры'
__author__ = 'SG'

import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *
from Autodesk.Revit.DB import ElementId, PartUtils, ViewOrientation3D, XYZ, FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, Line, Structure
import sys
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
# from Autodesk.Revit.ApplicationServices.Application import Create
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application
k = 304.8

t = Transaction(doc, 'Обернуть размеры')

t.Start()


sel = [doc.GetElement(elid) for elid in uidoc.Selection.GetElementIds()]


for el in sel:
    param1 = el.LookupParameter('Ширина воздуховода')
    param2 = el.LookupParameter('Высота воздуховода')
    if not param1:
        param1 = el.LookupParameter('Ширина')
        param2 = el.LookupParameter('Высота')
    width = param1.AsDouble()
    hight = param2.AsDouble()
    width, hight = hight, width
    param1.Set(width)
    param2.Set(hight)




t.Commit()
