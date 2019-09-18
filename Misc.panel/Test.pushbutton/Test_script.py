# # -*- coding: utf-8 -*-
# """"""
# __title__ = 'Test'
# __author__ = 'SG'

# import clr
# clr.AddReference('System.Core')
# from System.Collections.Generic import *
# from Autodesk.Revit.DB import ElementId, PartUtils, ViewOrientation3D, XYZ, FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, Line, Plane
# import sys
# from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
# # from Autodesk.Revit.ApplicationServices.Application import Create
# doc = __revit__.ActiveUIDocument.Document
# uidoc = __revit__.ActiveUIDocument
# app = __revit__.Application
# k = 304.8

# origin = XYZ.Zero
# normal = XYZ.BasisZ

# plane = Plane.Create(normal, origin)
# skplane = doc.FamilyCreate.NewSketchPlane(plane)

# import clr
# import math
# import numpy as np
# clr.AddReference('RevitAPI')
# clr.AddReference('RevitAPIUI')
# from Autodesk.Revit.DB import *

# doc = __revit__.ActiveUIDocument.Document
# app = __revit__.Application

# t = Transaction(doc, 'Create Line')

# t.Start()

# start = XYZ(0, 0, 0)
# end = XYZ(20, 20, 0)

# # normal =

# x = [1, 2, 3]
# y = [4, 5, 6]
# np.cross(x, y)

# plane = Plane.CreateByThreePoints(start, XYZ(1, 0, 0), end)
# # skplane = doc.FamilyCreate.NewSketchPlane(plane)
# sketchPlane = SketchPlane.Create(doc, plane)

# # Create line vertices

# # create NewLine()
# # line = app.Create.NewLine(lnStart, lnEnd, True)
# line = Line.CreateBound(start, end)

# modelLine = doc.Create.NewModelCurve(line, sketchPlane)

# # create NewModelCurve()
# # crv = doc.FamilyCreate.NewModelCurve(line, sketchPlane)

# t.Commit()








import clr
import math
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *

doc = __revit__.ActiveUIDocument.Document
app = __revit__.Application

t = Transaction(doc, 'Create Line')

t.Start()

start = XYZ(0, 0, 0)
end = XYZ(10, 10, 0)

line = Line.CreateBound(start, end)  #https://forums.autodesk.com/t5/revit-api-forum/how-to-crete-3d-modelcurves-to-avoid-exception-curve-must-be-in/td-p/8355936
direction = line.Direction
x, y, z = direction.X, direction.Y, direction.Z
# normal = XYZ(z - y, x - z, y - x)
normal = XYZ.BasisZ.CrossProduct(line.Direction)
plane = Plane.CreateByNormalAndOrigin(normal, start)
sketchPlane = SketchPlane.Create(doc, plane)
modelCurve = doc.Create.NewModelCurve(line, sketchPlane)



# plane = Plane.CreateByThreePoints(start, XYZ(1, 0, 0), end)
# # skplane = doc.FamilyCreate.NewSketchPlane(plane)
# sketchPlane = SketchPlane.Create(doc, plane)

# # Create line vertices

# # create NewLine()
# # line = app.Create.NewLine(lnStart, lnEnd, True)
# line = Line.CreateBound(start, end)

# modelLine = doc.Create.NewModelCurve(line, sketchPlane)

# # create NewModelCurve()
# # crv = doc.FamilyCreate.NewModelCurve(line, sketchPlane)

t.Commit()
