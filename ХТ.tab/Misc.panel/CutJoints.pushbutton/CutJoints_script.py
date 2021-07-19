# -*- coding: utf-8 -*-
""""""
__title__ = 'Вырезать\nв каркасе'
__author__ = 'SG'

import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *
from Autodesk.Revit.DB import PartUtils, ViewOrientation3D, XYZ, FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId, SolidSolidCutUtils
from Autodesk.Revit.DB import JoinGeometryUtils, BoundingBoxIntersectsFilter, Outline
import sys
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

sel = [doc.GetElement(elid) for elid in uidoc.Selection.GetElementIds()]

t = Transaction(doc, 'Вырезать в каркасе')
t.Start()

if __shiftclick__:
    from pyrevit import forms

    res = forms.alert('Следует выбрать вариант:',
                      options=["Соединить",
                               "Отменить соединение",
                               "Отмена"])

    walls = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements())
    floors = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType().ToElements())
    roofs = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Roofs).WhereElementIsNotElementType().ToElements())
    roofs = [i for i in roofs if (doc.GetElement(i.GetTypeId()).LookupParameter('Материал несущих конструкций') and
               'теплитель' in doc.GetElement(i.GetTypeId()).LookupParameter('Материал несущих конструкций').AsValueString())
               or
               (doc.GetElement(i.GetTypeId()).LookupParameter('м') and
               'теплитель' in doc.GetElement(i.GetTypeId()).LookupParameter('м').AsValueString())
               or
               'теплитель' in i.LookupParameter('Тип').AsValueString()]
    generic = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements())
    generic = [i for i in generic if (doc.GetElement(i.GetTypeId()).LookupParameter('Материал несущих конструкций') and
               'теплитель' in doc.GetElement(i.GetTypeId()).LookupParameter('Материал несущих конструкций').AsValueString())
               or
               (doc.GetElement(i.GetTypeId()).LookupParameter('м') and
               'теплитель' in doc.GetElement(i.GetTypeId()).LookupParameter('м').AsValueString())]

    for elem in walls + floors + generic + roofs:
        minpoint = elem.get_BoundingBox(doc.ActiveView).Min
        maxpoint = elem.get_BoundingBox(doc.ActiveView).Max
        els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).WherePasses(BoundingBoxIntersectsFilter(Outline(minpoint, maxpoint))).WhereElementIsNotElementType().ToElements()
        if res == "Соединить":
            for el in els:
                if not JoinGeometryUtils.AreElementsJoined(doc, el, elem):
                    JoinGeometryUtils.JoinGeometry(doc, el, elem)
        elif res == "Отменить соединение":
            for el in els:
                if JoinGeometryUtils.AreElementsJoined(doc, el, elem):
                    JoinGeometryUtils.UnjoinGeometry(doc, el, elem)

else:
    if len(sel) == 1:
        target = sel[0]
        cutting_ids = uidoc.Selection.PickObjects(ObjectType.Element, 'Выберите элементы, которые будут вырезаны из целевого')
        cutting_els = [doc.GetElement(elid.ElementId) for elid in cutting_ids]
        for el in cutting_els:
            if SolidSolidCutUtils.CanElementCutElement(target, el):
                SolidSolidCutUtils.AddCutBetweenSolids(doc, target, el)
    else:
        targets = sel
        cutting_id = uidoc.Selection.PickObject(ObjectType.Element, 'Выберите элемент, который будет вырезан из целевых')
        el = doc.GetElement(cutting_id.ElementId)
        for target in targets:
            if SolidSolidCutUtils.CanElementCutElement(target, el):
                SolidSolidCutUtils.AddCutBetweenSolids(doc, target, el, False)

    framings = []
    joints = []
    for el in sel:
        if el.Category.Name == 'Каркас несущий':
            # if el.Id not in framings:
            print(11)
            framings.append(el)
        elif el.Category.Name == 'Соединения несущих конструкций':
            joints.append(el)

    for framing in framings:
        for joint in joints:
            if not InstanceVoidCutUtils.InstanceVoidCutExists(framing, joint):
                InstanceVoidCutUtils.AddInstanceVoidCut(doc, framing, joint)

t.Commit()
