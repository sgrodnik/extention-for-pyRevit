# -*- coding: utf-8 -*-
""""""
__title__ = 'Test'
__author__ = 'SG'

import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *
from Autodesk.Revit.DB import ElementId, PartUtils, ViewOrientation3D, XYZ, FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId, SolidSolidCutUtils, InstanceVoidCutUtils
import sys
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
k = 304.8

sel = [doc.GetElement(elid) for elid in uidoc.Selection.GetElementIds()]

for part in sel:
	bbox = part.get_BoundingBox(doc.ActiveView)
	print(bbox.Max.Z * k)

# t = Transaction(doc, 'Test')
# t.Start()

# for elid in framings.Keys:
# 	framing = doc.GetElement(elid)
# 	for joint in joints:
# 		# print(framing)
# 		# print(joint)
# 		# SolidSolidCutUtils.AddCutBetweenSolids(doc, framing, joint)
# 		if not InstanceVoidCutUtils.InstanceVoidCutExists(framing, joint):
# 			InstanceVoidCutUtils.AddInstanceVoidCut(doc, framing, joint)


# # # sel = [doc.GetElement(elid) for elid indoc.Selection.GetElementIds()

# # parent_part_id = sel[0].GetSourceElementIds()[0].HostElementId
# # parent_id = doc.GetElement(parent_part_id).GetSourceElementIds()[0].HostElementId


# # # from pyrevit import script
# # # output = script.get_output()

# # els = PartUtils.GetAssociatedParts(doc, parent_id, False, True)
# # # for i in els:
# # # 	link = output.linkify(i, doc.GetElement(i))
# # # 	print('{}'.format(link))

# # uidoc.Selection.SetElementIds(els)

# # # print(PartUtils.HasAssociatedParts(doc, ElementId(269556)))

# t.Commit()
