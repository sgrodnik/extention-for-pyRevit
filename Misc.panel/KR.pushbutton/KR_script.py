# -*- coding: utf-8 -*-
""""""
__title__ = 'КР'
__author__ = 'SG'
import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId, XYZ, Structure

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

from pyrevit import script
output = script.get_output()

k = 304.8
k1 = 35.31466672149
k2 = 1000000 / k**2  # 10.763910416709722

t = Transaction(doc, 'КР')
t.Start()

els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).WhereElementIsNotElementType().ToElements()
for el in els:
    dimL = el.LookupParameter('Фактическая длина').AsDouble() * k
    el.LookupParameter('ADSK_Размер_Длина').Set(dimL / k)
    dimB = el.LookupParameter('ADSK_Размер_Высота').AsDouble() * k
    dimH = el.LookupParameter('ADSK_Размер_Ширина').AsDouble() * k
    name = ' {:.0f}×{:.0f}'.format(min([dimB, dimH]), max([dimB, dimH]))
    el.LookupParameter('ADSK_Марка').Set(name)
    if dimB >= 100 and dimH >= 100:
        name = 'Брус'
    else:
        name = 'Брусок' if dimB == dimH else 'Доска'
    doc.GetElement(el.GetTypeId()).LookupParameter('Описание').Set(name)
    zapas = 1.0 if 'Кле' in el.LookupParameter('Тип').AsValueString() else 1.1
    el.LookupParameter('ADSK_Количество').Set(dimL * dimB * dimH / 10**9 * zapas)

    None if el.LookupParameter('Этап').HasValue else el.LookupParameter('Этап').Set(999)

# Сортировка
els = [el for el in sorted(els, key=lambda x: x.LookupParameter('Фактическая длина').AsDouble(), reverse=True)]
els = [el for el in sorted(els, key=lambda x: float('.'.join(map(lambda x: '{:0>5}'.format(x), x.LookupParameter('ADSK_Марка').AsString().split('×')))), reverse=True)]
els = [el for el in sorted(els, key=lambda x: x.LookupParameter('Этап').AsDouble())]

# Простановка Позиции
i = 0
part = 0
name = 0
length = 0
for el in els:
    if 'Не СЭ' in el.LookupParameter('Тип').AsValueString():
        continue
    newpart = el.LookupParameter('Этап').AsDouble()
    newname = el.LookupParameter('ADSK_Марка').AsString()
    newlength = round(el.LookupParameter('Фактическая длина').AsDouble() * k, 0)
    if part != newpart or name != newname or length != newlength:
        i += 1
        part, name, length = newpart, newname, newlength
    el.LookupParameter('ADSK_Позиция').Set(str(i))


els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructConnections).WhereElementIsNotElementType().ToElements()
windows = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Windows).WhereElementIsNotElementType().ToElements())
doors = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Doors).WhereElementIsNotElementType().ToElements())

for el in els:
    None if el.LookupParameter('Этап').HasValue else el.LookupParameter('Этап').Set(999)
    el.LookupParameter('ADSK_Количество').Set(1)

walls = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements())
floors = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType().ToElements())
roofs = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Roofs).WhereElementIsNotElementType().ToElements())
fascias = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Fascia).WhereElementIsNotElementType().ToElements())
struct_connection_types = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructConnections).WhereElementIsElementType().ToElements()
for symbol in [i for i in struct_connection_types if 'Фейк' in i.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()]:
    None if 'Основа' in symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() else doc.Delete(symbol.Id)
struct_connection_types = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructConnections).WhereElementIsElementType().ToElements()
fake_original_symbol = [i for i in struct_connection_types if 'Фейк' in i.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()][0]

location = XYZ(0, 0, 0)
fake_symbols = {}
level = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()[0]
for exist_el in walls + floors + roofs + fascias:
    None if exist_el.LookupParameter('Этап').HasValue else exist_el.LookupParameter('Этап').Set(999)

    exist_symbol = doc.GetElement(exist_el.GetTypeId())
    description = exist_symbol.LookupParameter('Описание').AsString()
    description = description if description else exist_symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    for description in description.split('//'):
        if description not in fake_symbols:
            new_fake = fake_original_symbol.Duplicate('Фейк_ ' + description)
            fake_symbols[description] = new_fake
            new_fake.LookupParameter('Описание').Set(description.split('((')[0])
            key_note = exist_symbol.LookupParameter('Ключевая пометка').AsString()
            if not key_note:
                print('{} - Не указано значение параметра Ключевая пометка (единица измерения). Назначено значение поумолчанию "м³"'.format(output.linkify(exist_el.Id, description)))
                key_note = 'м³'
            new_fake.LookupParameter('Ключевая пометка').Set(key_note)
            cost = exist_symbol.LookupParameter('Стоимость').AsDouble()
            if not cost:
                cost = 777
            new_fake.LookupParameter('Стоимость').Set(cost)

        fake_symbol = fake_symbols[description]
        fake_el = doc.Create.NewFamilyInstance(location, fake_symbol, level, Structure.StructuralType.NonStructural)
        location += XYZ(0.01, -0.1, 0)
        key_note = exist_symbol.LookupParameter('Ключевая пометка').AsString()
        key_note = key_note if key_note else 'м³'
        coefficient = 1.1
        if '((' in description:
            coefficient = float(description.split('((')[-1].split('))')[0].replace(',', '.'))
        if key_note == 'м³':
            value = exist_el.LookupParameter('Объем').AsDouble() / k1 * coefficient
        elif key_note == 'м²':
            value = exist_el.LookupParameter('Площадь').AsDouble() / k2 * coefficient
        elif key_note == 'м':
            value = exist_el.LookupParameter('Длина').AsDouble() * k * coefficient / 1000
        fake_el.LookupParameter('ADSK_Количество').Set(value)
        # fake_el.LookupParameter('Наименование СМ').Set(description.split('((')[0])
        fake_el.LookupParameter('Комментарии').Set(str(exist_el.Id))
        fake_el.LookupParameter('Этап').Set(exist_el.LookupParameter('Этап').AsDouble())

t.Commit()
