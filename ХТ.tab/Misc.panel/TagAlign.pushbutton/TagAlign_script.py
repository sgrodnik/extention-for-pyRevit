# -*- coding: utf-8 -*-
"""Выравнивание марок"""
__title__ = 'Выровнять\nмарки'
__author__ = 'SG'
import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *
from Autodesk.Revit.DB import IndependentTag, XYZ, FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId
import sys
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

tags = [doc.GetElement(id) for id in uidoc.Selection.GetElementIds() if isinstance(doc.GetElement(id), IndependentTag)]

if not tags:
    sys.exit()

t = Transaction(doc, 'Выровнять марки')
t.Start()


# def align(tags):
#   if not tags:
#       return
#   elbow = {'x': None, 'y': None}
#   head = {'x': None, 'y': None}

#   elbows = {'x': [], 'y': []}
#   heads = {'x': [], 'y': []}
#   for tag in tags:
#       if tag.HasElbow:
#           elbows['x'].append(tag.LeaderElbow.X)
#           elbows['y'].append(tag.LeaderElbow.Y)
#       heads['x'].append(tag.TagHeadPosition.X)
#       heads['y'].append(tag.TagHeadPosition.Y)

#   elbow['x'] = sum(elbows['x']) / len(elbows['x'])
#   elbow['y'] = sum(elbows['y']) / len(elbows['y'])
#   head['x'] = sum(heads['x']) / len(heads['x'])
#   head['y'] = sum(heads['y']) / len(heads['y'])

#   for tag in tags:
#       tag.LeaderElbow = XYZ(elbow['x'], elbow['y'], 0)
#       tag.TagHeadPosition = XYZ(head['x'], head['y'], 0)

def align(tags):
    if not tags:
        return
    elbow, head = XYZ(), XYZ()

    for tag in tags:
        if tag.HasElbow:
            elbow = tag.LeaderElbow if elbow.IsZeroLength() else (elbow + tag.LeaderElbow) / 2
        head = tag.TagHeadPosition if head.IsZeroLength() else (head + tag.TagHeadPosition) / 2

    for tag in tags:
        tag.LeaderElbow = elbow
        tag.TagHeadPosition = head

lTags, rTags = [], []
for tag in tags:
    if tag.HasElbow: lTags.append(tag) if tag.LeaderElbow.X < tag.TagHeadPosition.X else rTags.append(tag)

align(rTags)
align(lTags)
# print(lTags[0].LeaderElbow)
# print(lTags[0].LeaderElbow / 2)


def move(tags, vector):
    for tag in tags:
        tag.LeaderElbow -= vector
        tag.TagHeadPosition -= vector


if rTags and lTags:
    vector = rTags[0].TagHeadPosition - lTags[0].TagHeadPosition
    move(rTags, vector)

t.Commit()

        # tag.LeaderElbow = rTags[0].LeaderElbow - (rTags[0].TagHeadPosition - lTags[0].TagHeadPosition)
# from pyrevit import script
# output = script.get_output()
#
# print(dir(output))
# 'self_destruct'

# print(dir(rTags[0].LeaderElbow))
# el.end = XYZ(1,2,0)
# el.elbow = XYZ(1,2,0)
# el.head = XYZ(1,2,0)
