# -*- coding: utf-8 -*-
"""Выбрать неисключенные части
Shift - выбрать исключенные части
Ctrl - выбрать все части"""
__title__ = 'Выбрать\nчасти'
__author__ = 'SG'

import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *
from Autodesk.Revit.DB import PartUtils, ViewOrientation3D, XYZ, FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId
import sys
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
# from pyrevit import script
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
import math

k = 304.8
k1 = 35.31466672149

sel = [doc.GetElement(elid) for elid in uidoc.Selection.GetElementIds() if doc.GetElement(elid).Name == 'Деталь']

# print(doc.GetElement(ElementId(313094)).GetSubelements())
# print(PartUtils.HasAssociatedParts(doc, ElementId(313094)))
# print(PartUtils.HasAssociatedParts(doc, ElementId(259996)))
# sys.exit()

if not sel:

    t = Transaction(doc, 'Выбрать части')
    t.Start()

    parts = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Parts).WhereElementIsNotElementType().ToElements()
    glue = 0
    seam_thickness = 2
    for part in parts:
        parent_part_id = part.GetSourceElementIds()[0].HostElementId
        parent_name = doc.GetElement(parent_part_id).Name
        if not part.Excluded and not PartUtils.HasAssociatedParts(doc, part.Id) and parent_name == 'Деталь':
            glue += part.LookupParameter('Высота').AsDouble() * k / 1000 * part.LookupParameter('Толщина').AsDouble() * k / 1000 * seam_thickness / 1000
            glue += (part.LookupParameter('Длина').AsDouble() * k + seam_thickness) / 1000 * part.LookupParameter('Толщина').AsDouble() * k / 1000 * seam_thickness / 1000
            bbox = part.get_BoundingBox(doc.ActiveView)
            part.LookupParameter('Ряд').Set(math.ceil(bbox.Max.Z * k / 252))

        if part.LookupParameter('Исходная категория').AsString() == 'Стены':
            part.LookupParameter('Наименование СМ').Set('Газобетонный {}блок толщиной {:.0f} мм'.format(
                'U-' if part.LookupParameter('Комментарии').AsString() == 'U-Блоки' else '',
                part.LookupParameter('Толщина').AsDouble() * k))
            volume = part.LookupParameter('Высота').AsDouble() * part.LookupParameter('Толщина').AsDouble() * part.LookupParameter('Длина').AsDouble() * k * k * k / 1000000000
            part.LookupParameter('ХТ Длина ОВ').Set(volume * 1.1)
        else:
            parent_part_id = part.GetSourceElementIds()[0].HostElementId
            parent = doc.GetElement(parent_part_id)
            name = doc.GetElement(parent.GetTypeId()).LookupParameter('Описание').AsString().replace('не вкл', '')
            part.LookupParameter('Наименование СМ').Set(name)
            # part.LookupParameter('Наименование СМ').Set('Бетон В20 W6 F150 по ГОСТ 26633-2012')
            part.LookupParameter('ХТ Длина ОВ').Set(part.LookupParameter('Объем').AsDouble() / k1 * 1.1)

        part.LookupParameter('ADSK_Единица измерения').Set('м³')


    els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()

    for el in els:
        if el.LookupParameter('Тип').AsValueString() == 'Клей для пеноблоков':
            el.LookupParameter('ХТ Длина ОВ').Set(glue * 1.1)
            el.LookupParameter('Наименование СМ').Set('Клей для пеноблоков')

        if doc.GetElement(el.GetTypeId()).LookupParameter('Описание').AsString() == 'Армопояс':
            el.LookupParameter('ХТ Длина ОВ').Set(el.LookupParameter('Объем').AsDouble() / k1 * 1.1)

        if doc.GetElement(el.GetTypeId()).LookupParameter('Описание').AsString() == 'Перемычки':
            el.LookupParameter('ХТ Длина ОВ').Set(el.LookupParameter('Объем').AsDouble() / k1 * 1.1)


    rebars = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
    for rebar in rebars:
        name = rebar.LookupParameter('Тип').AsValueString()
        rebar.LookupParameter('Наименование СМ').Set(name)
        if 'ø12' in name:
            to_kg = 0.888
        elif 'ø8' in name:
            to_kg = 0.395
        elif 'лента' in name:
            to_kg = 1.0
        rebar.LookupParameter('ХТ Длина ОВ').Set(rebar.LookupParameter('Полная длина стержня').AsDouble() * k / 1000 * 1.1 * to_kg)

    els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).WhereElementIsNotElementType().ToElements()
    els = [i for i in els if i.LookupParameter('Использование в конструкции').AsValueString() == 'Горизонтальный раскос']
    for el in els:
        el.LookupParameter('Наименование СМ').Set(doc.GetElement(el.GetTypeId()).LookupParameter('Описание').AsString())
        el.LookupParameter('ХТ Длина ОВ').Set(el.LookupParameter('Фактическая длина').AsDouble() * k / 1000 * 1.1)



    t.Commit()

else:
    parent_part_id = sel[0].GetSourceElementIds()[0].HostElementId
    parent_id = doc.GetElement(parent_part_id).GetSourceElementIds()[0].HostElementId

    t = Transaction(doc, 'Выбрать части')
    t.Start()

    # from pyrevit import script
    # output = script.get_output()

    el_ids = PartUtils.GetAssociatedParts(doc, parent_id, False, True)
    if __shiftclick__:
        el_ids = List[ElementId]([el_id for el_id in el_ids if doc.GetElement(el_id).Excluded])
    elif __forceddebugmode__:
        el_ids = List[ElementId]([el_id for el_id in el_ids])
    else:
        el_ids = List[ElementId]([el_id for el_id in el_ids if not doc.GetElement(el_id).Excluded])
    # for i in el_ids:
    #   link = output.linkify(i, doc.GetElement(i))
    #   print('{}'.format(link))

    uidoc.Selection.SetElementIds(el_ids)

    # print(PartUtils.HasAssociatedParts(doc, ElementId(269556)))

    t.Commit()
