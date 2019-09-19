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


location = XYZ(-10.8904389851532, 76.6607564383779, 0)
symbols = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsElementType().ToElements()
print(123)
print(list(symbols))
print(321)
symbol = list(filter(lambda x: x.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == 'Фейк 4', symbols))[0]
# levels = FilteredElementCollector(doc).OfClass(typeof(Level)).ToElements()
levels = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
level = levels[0]
doc.Create.NewFamilyInstance(location, symbol, level, Structure.StructuralType.NonStructural)


t.Commit()




