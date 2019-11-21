# -*- coding: utf-8 -*-
""""""
__title__ = 'Вырезать\nв каркасе'
__author__ = 'SG'

import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *
from Autodesk.Revit.DB import PartUtils, ViewOrientation3D, XYZ, FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId, SolidSolidCutUtils
import sys
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

sel = [doc.GetElement(elid) for elid in uidoc.Selection.GetElementIds()]

t = Transaction(doc, 'Вырезать в каркасе')
t.Start()

if len(sel) == 1:
    target = sel[0]
    cutting_ids = uidoc.Selection.PickObjects(ObjectType.Element, 'Выберите элеманты, которые будут вырезаны из целевого')
    cutting_els = [doc.GetElement(elid.ElementId) for elid in cutting_ids]
    for el in cutting_els:
        if SolidSolidCutUtils.CanElementCutElement(target, el):
            SolidSolidCutUtils.AddCutBetweenSolids(doc, target, el)
else:
    print('else')

# framings = []
# joints = []
# for el in sel:
#     if el.Category.Name == 'Каркас несущий':
#         # if el.Id not in framings:
#         print(11)
#         framings.append(el)
#     elif el.Category.Name == 'Соединения несущих конструкций':
#         joints.append(el)

# for framing in framings:
#     for joint in joints:
#         if not InstanceVoidCutUtils.InstanceVoidCutExists(framing, joint):
#             InstanceVoidCutUtils.AddInstanceVoidCut(doc, framing, joint)

t.Commit()
