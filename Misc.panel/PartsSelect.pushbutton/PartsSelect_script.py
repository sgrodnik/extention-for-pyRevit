# -*- coding: utf-8 -*-
""""""
__title__ = 'Выбрать\nчасти'
__author__ = 'SG'

import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *
from Autodesk.Revit.DB import PartUtils, ViewOrientation3D, XYZ, FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId
import sys
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

sel = [doc.GetElement(elid) for elid in uidoc.Selection.GetElementIds() if doc.GetElement(elid).Name == 'Деталь']

if not sel:
	sys.exit()

parent_part_id = sel[0].GetSourceElementIds()[0].HostElementId
parent_id = doc.GetElement(parent_part_id).GetSourceElementIds()[0].HostElementId

t = Transaction(doc, 'Выбрать части')
t.Start()

# from pyrevit import script
# output = script.get_output()

el_ids = PartUtils.GetAssociatedParts(doc, parent_id, False, True)
el_ids = List[ElementId]([el_id for el_id in el_ids if not doc.GetElement(el_id).Excluded])
# for i in el_ids:
# 	link = output.linkify(i, doc.GetElement(i))
# 	print('{}'.format(link))

uidoc.Selection.SetElementIds(el_ids)

# print(PartUtils.HasAssociatedParts(doc, ElementId(269556)))

t.Commit()
