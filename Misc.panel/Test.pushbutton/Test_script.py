# -*- coding: utf-8 -*-
""""""
__title__ = 'Test'
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

t = Transaction(doc, 'Test')

t.Start()



els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FlexDuctCurves).WhereElementIsElementType().ToElements()

type_ = list(filter(lambda x: x.LookupParameter('Имя типа').AsString() == 'SG_Круглый изолированный', els))[0]

sel = [doc.GetElement(elid) for elid in uidoc.Selection.GetElementIds()]

sel[0].FlexDuctType = type_

# for el in els:

# copy = els[0].Duplicate('asd')
# copy.Name = 'asd'

# print()

# for i in dir(els[0]):
# 	print(i)



t.Commit()

