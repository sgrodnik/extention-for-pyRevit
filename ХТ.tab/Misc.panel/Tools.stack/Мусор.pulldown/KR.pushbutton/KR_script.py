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
    dimB = el.LookupParameter('ADSK_Размер_Высота').AsDouble() * k
    dimH = el.LookupParameter('ADSK_Размер_Ширина').AsDouble() * k
    name = '{:.0f}×{:.0f}'.format(min([dimB, dimH]), max([dimB, dimH]))
    mark = name
    is_plita = el.LookupParameter('Семейство').AsValueString() == 'Плита'
    if is_plita:
        dimL = doc.GetElement(el.GetTypeId()).LookupParameter('h').AsDouble() * k
        dimB = el.LookupParameter('Ширина').AsDouble() * k
        dimH = el.LookupParameter('Фактическая длина').AsDouble() * k
        # name = ' {:.0f}×{:.0f}'.format(max([dimB, dimH]), min([dimB, dimH]))
        name = '{:.0f} мм'.format(dimL)
        mark = '{:.0f}×{:.0f}'.format(max([dimB, dimH]), min([dimB, dimH]))
    el.LookupParameter('ADSK_Размер_Длина').Set(dimL / k)
    el.LookupParameter('ADSK_Марка').Set(name)
    if dimB >= 100 and dimH >= 100:
        name, addition, addition2 = 'Брус', ' клееный из шпона (ЛВЛ)', ' ЛВЛ'
    else:
        name, addition, addition2 = ('Брусок', ' клееный из шпона (ЛВЛ)', ' ЛВЛ') if dimB == dimH else ('Доска', ' клееная из шпона (ЛВЛ)', ' ЛВЛ')
    is_glued = 'Кле' in el.LookupParameter('Тип').AsValueString()
    el.LookupParameter('Марка для поэлементной').Set(mark + (addition2 if is_glued else '' if is_plita else ' Сосна'))
    name += addition if is_glued else ' (сосна 1 сорта)'
    if is_plita:
        name = 'Плита типа ОСП-3, шлифованная, класса эмиссии Е1'
    doc.GetElement(el.GetTypeId()).LookupParameter('Описание').Set(name)
    zapas = 1.0 if is_glued else 1.1
    value = dimL * dimB * dimH / 10**9 * zapas
    if is_plita:
        value = dimB * dimH / 10**6 * zapas
    el.LookupParameter('Запас').Set('{0:n}'.format(zapas))
    el.LookupParameter('ADSK_Количество').Set(value)

    None if el.LookupParameter('Этап').HasValue else el.LookupParameter('Этап').Set(999)

derevo = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).WhereElementIsElementType().ToElements()
derevo = [i for i in derevo if 'Деревяха' in i.LookupParameter('Имя семейства').AsString()]
derevo = [el for el in sorted(derevo, key=lambda x: x.LookupParameter('b').AsDouble(), reverse=True)]
derevo = [el for el in sorted(derevo, key=lambda x: x.LookupParameter('h').AsDouble(), reverse=True)]
derevo = [el for el in sorted(derevo, key=lambda x: x.LookupParameter('b').AsDouble() * x.LookupParameter('h').AsDouble(), reverse=True)]
i1 = 1.01
i2 = 2.01
for el in derevo:
    if 'ЛВЛ' in el.LookupParameter('Описание').AsString():
        el.LookupParameter('Стоимость').Set(i1)
        i1 += 0.01
    if 'сосна' in el.LookupParameter('Описание').AsString():
        el.LookupParameter('Стоимость').Set(i2)
        i2 += 0.01

# Сортировка
els = [el for el in sorted(els, key=lambda x: x.LookupParameter('Фактическая длина').AsDouble(), reverse=True)]
els = [el for el in sorted(els, key=lambda x: float('.'.join(map(lambda x: '{:0>5}'.format(x), x.LookupParameter('ADSK_Марка').AsString().replace(' мм', '').split('×')))), reverse=True)]
els = [el for el in sorted(els, key=lambda x: x.LookupParameter('Этап').AsDouble())]
els = [el for el in sorted(els, key=lambda x: x.LookupParameter('Семейство').AsValueString() == 'Плита')]

# for i in els:
#     print('{:<10} {:<10} {:<10} '.format(
#         i.LookupParameter('Этап').AsDouble(),
#         i.LookupParameter('ADSK_Марка').AsString(),
#         i.LookupParameter('Фактическая длина').AsDouble() * k,
#     ))

# Простановка Позиции
i = 0
part = 0
name = 0
length = 0
done = {}
for el in els:
    if 'Не СЭ' in el.LookupParameter('Тип').AsValueString():
        continue
    newpart = el.LookupParameter('Этап').AsDouble()
    # newname = el.LookupParameter('ADSK_Марка').AsString()
    newname = el.LookupParameter('Марка для поэлементной').AsString()
    # is_plita = el.LookupParameter('Семейство').AsValueString() == 'Плита'
    # if is_plita:
    #     newname = el.LookupParameter('Марка для поэлементной').AsString()
    newlength = round(el.LookupParameter('Фактическая длина').AsDouble() * k, 0)
    if part != newpart or name != newname or length != newlength:
        if (newname, newlength) not in done:
            i += 1
            done[(newname, newlength)] = i
        part, name, length = newpart, newname, newlength
    # if (newname, newlength) not in done:
    #     el.LookupParameter('ADSK_Позиция').Set(str(i))
    position = done[(newname, newlength)]
    # print(position)
    el.LookupParameter('ADSK_Позиция').Set(str(position))
    # else:
    #     el.LookupParameter('ADSK_Позиция').Set(str(i))


els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructConnections).WhereElementIsNotElementType().ToElements()
windows = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Windows).WhereElementIsNotElementType().ToElements())
doors = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Doors).WhereElementIsNotElementType().ToElements())

for el in windows + doors:
    el.LookupParameter('Марка для поэлементной').Set('Не учитывать')

for el in els:
    None if el.LookupParameter('Этап').HasValue else el.LookupParameter('Этап').Set(999)
    el.LookupParameter('ADSK_Количество').Set(1)

walls = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements())
floors = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType().ToElements())
roofs = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Roofs).WhereElementIsNotElementType().ToElements())
fascias = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Fascia).WhereElementIsNotElementType().ToElements())
generic = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements())
struct_connection_types = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructConnections).WhereElementIsElementType().ToElements()
for symbol in [i for i in struct_connection_types if 'Фейк' in i.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()]:
    None if 'Основа' in symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() else doc.Delete(symbol.Id)
struct_connection_types = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructConnections).WhereElementIsElementType().ToElements()
fake_original_symbol = [i for i in struct_connection_types if 'Фейк' in i.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()][0]

location = XYZ(0, 0, 0)
fake_symbols = {}
level = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()[0]
for exist_el in walls + floors + roofs + fascias + generic:
    None if exist_el.LookupParameter('Этап').HasValue else exist_el.LookupParameter('Этап').Set(999)

    exist_symbol = doc.GetElement(exist_el.GetTypeId())
    description = exist_symbol.LookupParameter('Описание').AsString()
    if 'е учит' in description:
        continue
    description = description if description else exist_symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    if 'Формула' in description:
        description = exist_symbol.LookupParameter('Формула').AsString().replace('\r', '')
        exist_symbol.LookupParameter('Формула строкой').Set(description.replace('\n', ' // '))
    values = []
    for description in description.split('\n'):
        if description not in fake_symbols:
            new_fake = fake_original_symbol.Duplicate('Фейк_ ' + description)
            fake_symbols[description] = new_fake
            new_fake.LookupParameter('Описание').Set(description.split('%')[0])
            key_note = exist_symbol.LookupParameter('Ключевая пометка').AsString()
            if '%3' in description:
                key_note = description.split('%3')[1]
            if not key_note:
                print('{} - Не указано значение параметра Ключевая пометка (единица измерения). Назначено значение поумолчанию "м³"'.format(output.linkify(exist_el.Id, description)))
                key_note = 'м³'
            new_fake.LookupParameter('Ключевая пометка').Set(key_note)
            cost = exist_symbol.LookupParameter('Стоимость').AsDouble()
            if '%0' in description:
                cost = float(description.split('%0')[1].replace(',', '.'))
            if not cost:
                cost = 777
            new_fake.LookupParameter('Стоимость').Set(cost)

        fake_symbol = fake_symbols[description]
        fake_el = doc.Create.NewFamilyInstance(location, fake_symbol, level, Structure.StructuralType.NonStructural)
        location += XYZ(0.01, -0.1, 0)
        key_note = exist_symbol.LookupParameter('Ключевая пометка').AsString()
        key_note = key_note if key_note else 'м³'
        coefficient = 1
        zapas = 1.1
        if '%4' in description:
            coefficient = eval(description.split('%4')[1].replace(',', '.'))
            if not coefficient:
                raise Exception('Вероятно, забыта десятичная точка в формуле')
        if '%5' in description:
            zapas = eval(description.split('%5')[1].replace(',', '.'))
        fake_el.LookupParameter('Запас').Set('{0:n}'.format(zapas))
        marka = ''
        if '%1' in description:
            marka = description.split('%1')[1]
        fake_el.LookupParameter('ADSK_Марка').Set(marka)
        if key_note == 'м³':
            value = exist_el.LookupParameter('Объем').AsDouble() / k1 * coefficient * zapas
        elif key_note == 'м²':
            value = exist_el.LookupParameter('Площадь').AsDouble() / k2 * coefficient * zapas
        elif key_note == 'м':
            value = exist_el.LookupParameter('Длина').AsDouble() * k * coefficient * zapas / 1000
        fake_el.LookupParameter('ADSK_Количество').Set(value)
        values.append('{} {}: {:.3f}'.format(new_fake.LookupParameter('Описание').AsString(), marka, value))
        # fake_el.LookupParameter('Наименование СМ').Set(description.split('((')[0])
        fake_el.LookupParameter('Комментарии').Set(str(exist_el.Id))
        fake_el.LookupParameter('Этап').Set(exist_el.LookupParameter('Этап').AsDouble())
    exist_el.LookupParameter('Наименование СМ').Set('\n'.join(values))

t.Commit()
