# -*- coding: utf-8 -*-
""""""
__title__ = 'Расчёт\nАР'
__author__ = 'SG'

import clr
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

t = Transaction(doc, 'Расчёт АР')

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
lightingFixtures = FilteredElementCollector(doc).OfCategory(
    BuiltInCategory.OST_LightingFixtures).WhereElementIsNotElementType().ToElements()
ductTerminals = FilteredElementCollector(doc).OfCategory(
    BuiltInCategory.OST_DuctTerminal).WhereElementIsNotElementType().ToElements()

# for room in rooms:
for el in doors:
    poms = []
    if el.FromRoom[phase]:
        number = el.FromRoom[phase].LookupParameter('Номер').AsString()
        if number == '7104': number += 'а'
        poms.append(number)
    if el.ToRoom[phase]:
        number = el.ToRoom[phase].LookupParameter('Номер').AsString()
        if number == '7104': number += 'а'
        poms.append(number)
    el.LookupParameter('Пом').Set(' '.join(sorted(poms)))
    el.LookupParameter('Количество').Set(0)

for el in windows:
    poms = []
    if True:
        if el.FromRoom[phase]:
            number = el.FromRoom[phase].LookupParameter('Номер').AsString()
            if number == '7104': number += 'а'
            poms.append(number)
        if el.ToRoom[phase]:
            number = el.ToRoom[phase].LookupParameter('Номер').AsString()
            if number == '7104': number += 'а'
            poms.append(number)
    el.LookupParameter('Пом').Set(' '.join(sorted(poms)))
    el.LookupParameter('Количество').Set(0)

for el in lightingFixtures:
    # print(el.Id)
    el.LookupParameter('Количество').Set(0)
    if el.Room[phase]:
        number = el.Room[phase].LookupParameter('Номер').AsString()
        if number == '7104': number += 'а'
        el.LookupParameter('Пом').Set(number)

for el in ductTerminals:
    # print(el.Id)
    el.LookupParameter('Количество').Set(0)
    if el.Room[phase]:
        number = el.Room[phase].LookupParameter('Номер').AsString()
        if number == '7104': number += 'а'
        el.LookupParameter('Пом').Set(number)


els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
for el in filter(lambda x: 'Фейк' in x.LookupParameter('Тип').AsValueString(), els):
    doc.Delete(el.Id)

levels = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
level = levels[0]
location = XYZ(0, 0, 0)

symbols = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsElementType().ToElements()

def get_symbol(name):
    global symbols
    symbol = list(filter(lambda x: name[1:] in x.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString(), symbols))
    if symbol:
        return symbol[0]
    else:
        t.RollBack()
        raise Exception('Создайте тип для Фейк ' + name)

for room in rooms:
    floor = room.LookupParameter('Отделка пола').AsString()
    ceiling = room.LookupParameter('Отделка потолка').AsString()
    wall = room.LookupParameter('Отделка стен').AsString()

    area = room.LookupParameter('Площадь').AsDouble() / kk
    perimeter = room.LookupParameter('Периметр').AsDouble() * k / 1000
    sposob = room.LookupParameter('Способ расчета площади').AsDouble()
    visota = room.LookupParameter('Полная высота').AsDouble() * k / 1000
    area_doors = room.LookupParameter('ADSK_Площадь проемов').AsDouble() / kk
    # Линолеум Tarkett IQ Monolit (КМ2)
    # Керамическая плитка, нескользящая

    el = doc.Create.NewFamilyInstance(location, get_symbol(floor), level, Structure.StructuralType.NonStructural)
    if sposob > 0:
        result = area
    else:
        result = area + perimeter * 0.1
    # if room.LookupParameter('Номер').AsString() == '4025':
    #   print(area)
    #   print(perimeter)
    #   print(sposob)
    #   print(result)
    #   print(visota)
    #   print(area_doors)
    el.LookupParameter('Количество').Set(result)
    el.LookupParameter('ADSK_Этаж').Set(room.LookupParameter('Номер').AsString())
    number = room.LookupParameter('Номер').AsString()
    if number == '7104': number += 'а'
    el.LookupParameter('Пом').Set(number)

    el = doc.Create.NewFamilyInstance(location + XYZ(1, 0, 0), get_symbol(ceiling), level, Structure.StructuralType.NonStructural)
    el.LookupParameter('Количество').Set(area)
    el.LookupParameter('ADSK_Этаж').Set(room.LookupParameter('Номер').AsString())
    number = room.LookupParameter('Номер').AsString()
    if number == '7104': number += 'а'
    el.LookupParameter('Пом').Set(number)

    el = doc.Create.NewFamilyInstance(location + XYZ(2, 0, 0), get_symbol(wall), level, Structure.StructuralType.NonStructural)
    walls_area = perimeter * visota - area_doors
    # if room.LookupParameter('Номер').AsString() == '7006':
    #     print(glass)
    #     print(rentgen)
    el.LookupParameter('Количество').Set(walls_area)
    el.LookupParameter('ADSK_Этаж').Set(room.LookupParameter('Номер').AsString())
    number = room.LookupParameter('Номер').AsString()
    if number == '7104': number += 'а'
    el.LookupParameter('Пом').Set(number)

    glass = room.LookupParameter('Длина стекла').AsDouble() * k / 1000 * 2.2
    rentgen = room.LookupParameter('Длина рентгена').AsDouble() * k / 1000 * 3.8
    if glass:
        room.LookupParameter('Примечание').Set('Sостекл={:.1f}м²; Sоцинк={:.1f}м²'.format(glass, walls_area - glass).replace('.', ','))

        el = doc.Create.NewFamilyInstance(location + XYZ(4, 0, 0), get_symbol('Остекление'), level, Structure.StructuralType.NonStructural)
        el.LookupParameter('Количество').Set(glass)
        el.LookupParameter('ADSK_Этаж').Set(room.LookupParameter('Номер').AsString())
        number = room.LookupParameter('Номер').AsString()
        if number == '7104': number += 'а'
        el.LookupParameter('Пом').Set(number)

        el = doc.Create.NewFamilyInstance(location + XYZ(5, 0, 0), get_symbol('Оцинковка'), level, Structure.StructuralType.NonStructural)
        el.LookupParameter('Количество').Set(walls_area - glass)
        el.LookupParameter('ADSK_Этаж').Set(room.LookupParameter('Номер').AsString())
        number = room.LookupParameter('Номер').AsString()
        if number == '7104': number += 'а'
        el.LookupParameter('Пом').Set(number)

    if rentgen:
        room.LookupParameter('Примечание').Set('Sрентген={:.1f}м²'.format(rentgen).replace('.', ','))

        el = doc.Create.NewFamilyInstance(location + XYZ(3, 0, 0), get_symbol('Рентген'), level, Structure.StructuralType.NonStructural)
        # print(rentgen)
        el.LookupParameter('Количество').Set(rentgen)
        el.LookupParameter('ADSK_Этаж').Set(room.LookupParameter('Номер').AsString())
        number = room.LookupParameter('Номер').AsString()
        if number == '7104': number += 'а'
        el.LookupParameter('Пом').Set(number)





    location += XYZ(0, -0.05, 0)





t.Commit()
