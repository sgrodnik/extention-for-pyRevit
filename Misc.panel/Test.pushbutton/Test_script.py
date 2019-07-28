# -*- coding: utf-8 -*-
""""""
__title__ = 'Test'
__author__ = 'SG'
import os
import shutil
import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *
from Autodesk.Revit.DB import ViewOrientation3D, XYZ, FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId
import sys
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

path = doc.PathName + '.txt'

t = Transaction(doc, 'Test')
t.Start()

if __shiftclick__:
	with open(path, 'w') as file:
		zyx = doc.ActiveView.GetOrientation().EyePosition
		file.write('{:.16f} {:.16f} {:.16f}\n'.format(zyx[0], zyx[1], zyx[2]))
		zyx = doc.ActiveView.GetOrientation().ForwardDirection
		file.write('{:.16f} {:.16f} {:.16f}\n'.format(zyx[0], zyx[1], zyx[2]))
		zyx = doc.ActiveView.GetOrientation().UpDirection
		file.write('{:.16f} {:.16f} {:.16f}\n'.format(zyx[0], zyx[1], zyx[2]))
else:
	with open(path, 'r') as file:
		# print(file.read())
		file = file.read()
		line = file.split('\n')[0]
		eyePosition = XYZ(float(line.split()[0]), float(line.split()[1]), float(line.split()[2]))
		line = file.split('\n')[1]
		forwardDirection = XYZ(float(line.split()[0]), float(line.split()[1]), float(line.split()[2]))
		line = file.split('\n')[2]
		upDirection = XYZ(float(line.split()[0]), float(line.split()[1]), float(line.split()[2]))

		vo = ViewOrientation3D(eyePosition, upDirection, forwardDirection)

		doc.ActiveView.SetOrientation(vo)

t.Commit()
