# -*- coding: utf-8 -*-
""""""
__title__ = 'Обрезка\n3D вида'
__author__ = 'SG'
import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *
from Autodesk.Revit.DB import Point, BoundingBoxXYZ, XYZ, FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId
from Autodesk.Revit.UI.Selection import ObjectType, ObjectSnapTypes
import sys
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

sectionBoxes = FilteredElementCollector(doc)\
    .OfCategory(BuiltInCategory.OST_SectionBox)\
    .WhereElementIsNotElementType().ToElements()

box = doc.ActiveView.GetSectionBox()
min = box.Min
max = box.Max
origin = box.Transform.Origin
# print(min)
# print(max)
# print(origin)

try:
    target = uidoc.Selection.PickObject(ObjectType.Face, 'Выберите целевую точку')
    # target = uidoc.Selection.PickPoint('Выберите целевую точку')
except:  # Exceptions.OperationCanceledException:
    sys.exit()

el = doc.GetElement(target.ElementId)
face = el.GetGeometryObjectFromReference(target)
target = face.Origin

print(face.Origin)

# uidoc.Selection.SetElementIds(face.Id)

# target = XYZ(0, 0, 3)
abs = origin.Z + max.Z
delta = target.Z - abs

box.Max = XYZ(max.X, max.Y, max.Z + delta)

# print(face.Origin)
# print(face.GetBoundingBox())
# print(dir(face))


# box.Max -= XYZ(0, 0, 1)


t = Transaction(doc, 'Обрезка 3D вида')
t.Start()

# Point.Create(XYZ(0, 0, 10))

# doc.ActiveView.SetSectionBox(box)

t.Commit()

# Document doc = uidoc.Document;

# Reference r = uidoc.Selection.PickObject(
#   ObjectType.Face,
#   "Please select a planar face to define work plane" );

# Element e = doc.get_Element( r.ElementId );

# if( null != e )
# {
#   PlanarFace face
#     = e.GetGeometryObjectFromReference( r )
#       as PlanarFace;