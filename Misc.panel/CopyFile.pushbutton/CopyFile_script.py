# -*- coding: utf-8 -*-
""""""
__title__ = 'CopyFile'
__author__ = 'SG'
import os
import shutil
import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *
from Autodesk.Revit.DB import Mechanical, FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId
import sys
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

els = [doc.GetElement(id) for id in uidoc.Selection.GetElementIds()]

t = Transaction(doc, 'CopyFile')
t.Start()

src = doc.PathName
path = os.path.dirname(src)
dst = path + '123.rvt'
# print(path + '123.rvt')
for i in os.listdir(r'D:\G\1\Берег'):
	if i.split('.')[-1] == 'lnk':
		print i
# shutil.copy2(src, dst)
# shutil.copyfile(r'/home/py/mouse.txt', r'/home/py/new-mouse.txt')

t.Commit()
