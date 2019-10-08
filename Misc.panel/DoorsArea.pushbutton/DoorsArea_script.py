# -*- coding: utf-8 -*-
"""Суммирует площади окон и дверей помещения и записывает сумму в 'ADSK_Площадь проемов'"""
__title__ = 'Площадь\nпроёмов'
__author__ = 'SG'

import clr
import re
clr.AddReference('System.Core')
from System.Collections.Generic import *
from Autodesk.Revit.DB import ElementId, PartUtils, ViewOrientation3D, XYZ, FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, Line, Structure
import sys
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
# from Autodesk.Revit.ApplicationServices.Application import Create
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application
k = 304.8
kk = 1 / 0.3048**2


def natural_sorted(list, key=lambda s: s):
    """
    Sort the list into natural alphanumeric order.
    """
    def get_alphanum_key_func(key):
        convert = lambda text: int(text) if text.isdigit() else text
        return lambda s: [convert(c) for c in re.split('([0-9]+)', key(s))]
    sort_key = get_alphanum_key_func(key)
    return sorted(list, key=sort_key)


t = Transaction(doc, 'Площадь проёмов')

t.Start()

phases = doc.Phases
phase = phases[phases.Size - 1]
sel = [doc.GetElement(elid) for elid in uidoc.Selection.GetElementIds()]


rooms = FilteredElementCollector(doc).OfCategory(
    BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()
doors = FilteredElementCollector(doc).OfCategory(
    BuiltInCategory.OST_Doors).WhereElementIsNotElementType().ToElements()
windows = FilteredElementCollector(doc).OfCategory(
    BuiltInCategory.OST_Windows).WhereElementIsNotElementType().ToElements()
data = []
for room in rooms:
    # if room.LookupParameter('Номер').AsString() == '4050':
    # if room.LookupParameter('Номер').AsString()[0] == '7':
    if True:
        area = 0
        door_counter = 0
        door_types = []
        for door in doors:
            if door.LookupParameter('Тип').AsValueString() != 'ADSK_Дверь_Проем':
                if door.FromRoom[phase]:
                    if door.FromRoom[phase].LookupParameter('Номер').AsString() == room.LookupParameter('Номер').AsString():
                        door_type = doc.GetElement(door.GetTypeId())
                        area += door_type.LookupParameter('Высота').AsDouble(
                        ) * door_type.LookupParameter('Ширина').AsDouble()
                        # print(door.Id)
                        # print(door_type.LookupParameter('Высота').AsDouble() * door_type.LookupParameter('Ширина').AsDouble())
                        door_counter += 1
                        door_types.append(door.LookupParameter('Тип').AsValueString())
                if door.ToRoom[phase]:
                    if door.ToRoom[phase].LookupParameter('Номер').AsString() == room.LookupParameter('Номер').AsString():
                        door_type = doc.GetElement(door.GetTypeId())
                        area += door_type.LookupParameter('Высота').AsDouble(
                        ) * door_type.LookupParameter('Ширина').AsDouble()
                        # print(door.Id)
                        # print(door_type.LookupParameter('Высота').AsDouble() * door_type.LookupParameter('Ширина').AsDouble())
                        door_counter += 1
                        door_types.append(door.LookupParameter('Тип').AsValueString())
        window_counter = 0
        window_types = []
        for window in windows:
            if 'Л-' not in window.LookupParameter('Тип').AsValueString():
                if window.FromRoom[phase]:
                    if window.FromRoom[phase].LookupParameter('Номер').AsString() == room.LookupParameter('Номер').AsString():
                        window_type = doc.GetElement(window.GetTypeId())
                        area += window.LookupParameter('Высота').AsDouble(
                        ) * window.LookupParameter('Ширина').AsDouble()
                        # print(window.Id)
                        # print(window.LookupParameter('Высота').AsDouble() * window.LookupParameter('Ширина').AsDouble())
                        window_counter += 1
                        window_types.append(window.LookupParameter('Тип').AsValueString())
                if window.ToRoom[phase]:
                    if window.ToRoom[phase].LookupParameter('Номер').AsString() == room.LookupParameter('Номер').AsString():
                        window_type = doc.GetElement(window.GetTypeId())
                        area += window.LookupParameter('Высота').AsDouble(
                        ) * window.LookupParameter('Ширина').AsDouble()
                        # print(window.Id)
                        # print(window.LookupParameter('Высота').AsDouble() * window.LookupParameter('Ширина').AsDouble())
                        window_counter += 1
                        window_types.append(window.LookupParameter('Тип').AsValueString())
        room.LookupParameter('ADSK_Площадь проемов').Set(area)
    # door_types = list(set(door_types)).sort()
    # door_types.sort() if door_types else None
    # window_types = list(set(window_types)).sort()
    # window_types.sort() if window_types else None
    door_types = list(set(door_types))
    door_types = natural_sorted(door_types)
    door_types = '; '.join(door_types)
    window_types = list(set(window_types))
    window_types = natural_sorted(window_types)
    window_types = '; '.join(window_types)
    data.append([room.LookupParameter('Номер').AsString(), door_counter if door_counter else '', window_counter if window_counter else '',
                 area / kk, door_types, window_types])
# data = [
# ['row1', 'data', 'data', 80 ],
# ['row2', 'data', 'data', 45 ],
# ]
data.sort(key=lambda x: x[0])
from pyrevit import script
output = script.get_output()
output.print_table(
    table_data=data,
    # title="Example Table",
    columns=["№", "Двери", "Окна", "Площадь проемов", "Типы дверей", "Типы окон"],
    formats=['', '', '', '{:.2f}', '', ''],
    # last_line_style='color:red;'
)


t.Commit()
