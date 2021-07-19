# -*- coding: utf-8 -*-
"""(Добавить описание и инструкцию)"""
__title__ = 'Расчет\nMVS'
__author__ = 'SG'

import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *
from Autodesk.Revit.DB import ElementId, PartUtils, ViewOrientation3D, XYZ, FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, Line, Point, SketchPlane, Plane, Structure
import sys
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
# import math
from math import ceil

from pyrevit import script
output = script.get_output()

els = FilteredElementCollector(doc).OfCategory(
    BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().ToElements()

for el in els:  # Проверки наличия панелей и количества КабелейСигнальных
    if el.MEPModel.ElectricalSystems:
        # arr = [cir for cir in list(el.MEPModel.ElectricalSystems) if cir.BaseEquipment.Id != el.Id]

        arr = []
        for cir in list(el.MEPModel.ElectricalSystems):
            if not cir.BaseEquipment:
                name = el.LookupParameter('Тип').AsValueString()
                cir_number = cir.LookupParameter('Номер цепи').AsString()
                print('Не указана панель: {}, номер цепи {}'.format(output.linkify(el.Id, name), cir_number))
                raise Exception('Не указана панель')
            if cir.BaseEquipment.Id != el.Id:
                arr.append(cir)
        if arr:
            len1 = len(arr)
            # print(el.Id)
            # print(doc.GetElement(el.GetTypeId()))
            len2 = len(doc.GetElement(el.GetTypeId()).LookupParameter('КабельСигнальный').AsString().split('/'))
            name = el.LookupParameter('Тип').AsValueString()
            if len1 > len2:
                print('Ошибка: {}: подключений: {} КабельСигнальный: {}'.format(output.linkify(el.Id, name), len1, len2))
            elif len1 < len2:
                print('Предупреждение: {}: подключений: {} КабельСигнальный: {}'.format(output.linkify(el.Id, name), len1, len2))
                # print('{}'.format(output.linkify(mydict[name], name)))


cirs = FilteredElementCollector(doc).OfCategory(
    BuiltInCategory.OST_ElectricalCircuit).WhereElementIsNotElementType().ToElements()

firstElementsOfCirs = [list(cir.Elements)[0].Id for cir in cirs]

pairs = [(key, cir) for key, cir in zip(firstElementsOfCirs, cirs)]

pairs.sort(key=lambda x: x[0])

sortedCirs = [p[1] for p in pairs]

cirs = [i for i in sortedCirs]

panels = []  # Далее формируем список панелей цепей
elements = []  # Далее формируем двумерный список элементов цепей
errors = []
for cir in (cirs):
    panels.append(cir.BaseEquipment)  # Возвращает элемент
    if not cir.BaseEquipment:
        errors.append(cir)
    elements.append(cir.Elements)  # Возвращает список(!) элементов

if errors:
    for cir in errors:
        print(cir.Id)
    raise Exception('Обнаружены неподключенные цепи. Не указана Панель')

namesOfTypesOfPanels = []  # Далее формируем список имён типов панелей
for panel in panels:
    type = doc.GetElement(panel.GetTypeId())
    namesOfTypesOfPanels.append(type.get_Parameter(
        BuiltInParameter.SYMBOL_NAME_PARAM).AsString())

wires = []  # Далее формируем список кабелей цепей
namesOfTypesOfElements = []  # Далее формируем список имён типов элементов
rooms = []
for i, element in enumerate(elements):
    wire = []
    name = []
    subRooms = []
    for el in element:  # листаем вложенный список
        type = doc.GetElement(el.GetTypeId())
        # В зависимости от того, силовая ли цепь или нет, берём строку с кабелями типа
        if panels[i].Name == 'Силовой шкаф':
            wireOfType = type.LookupParameter('КабельСиловой').AsString()
        else:
            wireOfType = type.LookupParameter('КабельСигнальный').AsString()
        wire.append(wireOfType)
        name.append(type.get_Parameter(
            BuiltInParameter.SYMBOL_NAME_PARAM).AsString().replace(' значок выкл', ''))
        if not el.LookupParameter('Помещение').AsString():
            # print(doc.GetElement(el.Id).LookupParameter('Тип').AsValueString())
            print('{} Следует заполнить параметр Помещение'.format(output.linkify(el.Id, el.LookupParameter('Тип').AsValueString())))
            raise Exception('Помещение')
        subRooms.append(el.LookupParameter('Помещение').AsString())
    if len(wire) == 1:  # если длина вложенного списка == 1 (то есть в цепи один Элемент, а не шлейф), то
        wires.append(wire[0])  # берём значение единственного элемента
        namesOfTypesOfElements.append(name[0])
    else:  # иначе длина вложенного спискае больше 1, то
        string = wire[0]  # для начала формируем строку из первого значения
        for w in wire:
            if string != w:  # если значения не одинаковые, то значит, что в шлейф соединены Оборуды с разными значениями параметра КабельСигнальный, это не правильно
                string = 'Ошибка 1'  # формируем ошибку
        # Теперь поступим аналогично с именами типов: для начала формируем строку из первого значения
        string2 = name[0]
        for n in name:
            # если значения не одинаковые, то        # ВЕРОЯТНО, ЭТО НАДО ИСПРАВИТЬ (чтобы делать шлейфы с разными оборудами при одинаковом кабеле)
            if string2 != n:
                string2 = 'Ошибка 2'  # формируем ошибку
        wires.append(string)  # забираем результат
        namesOfTypesOfElements.append(string2)
    # wire_as_slashed_sring = wires[-1]
    # print(wire_as_slashed_sring)
    rooms.append(subRooms[0])

t = Transaction(doc, 'Расчет MVS')
t.Start()

for i, cir in enumerate(cirs):
    # "Имя типа" Панели прописываем в "Имя панели"
    panels[i].LookupParameter('Имя панели').Set(namesOfTypesOfPanels[i])
    # "Имя типа" + "Помещение" Оборуды прописываем в "Имя нагрузки" Цепи
    cir.LookupParameter('Имя нагрузки').Set(namesOfTypesOfElements[i])
    cir.LookupParameter('Помещение').Set(rooms[i])


# Чистим список: если типу Оборуды не прописан КабельСигнальный то будет None; заменим None на '-'
for i, wire in enumerate(wires):
    if wire is None:
        wires.pop(i)
        wires.insert(i, '-')

#for i in sorted(wires):
#    print(i)

"""
numberOfWires = [] # Позднее нам потребуется знать количество разделителей
for wire in wires:
    if wire.find('/') == -1:
        numberOfWires.append(0)       --------УДАЛИТЬ ЭТОТ КУСОК-----------
    else:
        numberOfWires.append(wire.count('/'))"""



# counter = 1
# doneList = []
# dividedWires = [i for i in wires]

# # for i in dividedWires:
# #     print(i)
# # print('0-----------------0')

# # print(11)
# # print(len(dividedWires))

# # Хитрый цикл, в котором стринги кабелей с разделителями заменяем отдельными кабелями
# for i, wire in enumerate(dividedWires):
#     if wire.count('/') > 0:
#         # print(wire)
#         if i not in doneList:
#             for j, w in enumerate(wire.split('/')):
#                 # print(w)
#                 try:
#                     dividedWires.pop(i + j)
#                 except:
#                     a = 'похер'  # pass
#                 dividedWires.insert(i + j, w)
#                 doneList.append(i + j)
#                 # print('{} len = {}'.format(counter, len(dividedWires)))
#                 counter += 1
# # for i in dividedWires:
# #     print(i)
# # print('1-----------------1')

# # print(22)
# # print(len(dividedWires))

cirs_grouped_by_element = {}
for cir in cirs:
    element = list(cir.Elements)[0]
    if element.Id not in cirs_grouped_by_element:
        cirs_grouped_by_element[element.Id] = []
    for connector in element.MEPModel.ConnectorManager.Connectors:
        all_refs = list(connector.AllRefs)
        if all_refs:
            if all_refs[0].Owner.Id == cir.Id:
                cirs_grouped_by_element[element.Id].append((connector.Id, cir))

cirs_and_their_wires = {}
for el_id in cirs_grouped_by_element:
    el = doc.GetElement(el_id)
    wire_as_slashed_sring = doc.GetElement(el.GetTypeId()).LookupParameter('КабельСигнальный').AsString()
    wire_list = wire_as_slashed_sring.split('/')
    cirs_group = cirs_grouped_by_element[el_id]
    counter = 0
    for number, cir in sorted(cirs_group, key=lambda (n, c): n):
        cir_wire = wire_list[counter]
        counter += 1
        cirs_and_their_wires[cir.Id] = cir_wire


dividedWires = []
for cir in cirs:
    dividedWires.append(cirs_and_their_wires[cir.Id])


#-------------Смена режима траектории--------------------

# specialElements = ['Рабочая станция', 'Конвертер 3G-SDI/HDMI', 'Розетка HDMI', 'Монитор управления',
#                    'Консоль (розетка HDMI)', 'Переходник DVI-D/HDMI', 'Розетка HDMI', 'Масштабатор', 'Монитор управления', 'USB2', 'Блок розеток HDMI, DVI (HDMI), SDI, CVBS, + DVI (HDMI), VGA']  # список оборудования, от которого прокладка ведётся по кратчайшему пути
# specialPanels = ['MiniPC', 'Монитор обзорный', 'Конвертер HDMI/SDI', 'MiniPC', 'Конвертер HDMI/SDI', 'Розетка HDMI',
#                  'Масштабатор', 'Конвертер HDMI/SDI', 'USB1', 'Compact A3', 'БААВС']  # список панелей, к которым прокладка ведётся по кратчайшему пути

from Autodesk.Revit.DB import Electrical

IN = 1
if IN:
    modes = []
    for index, cir in enumerate(cirs):
        if dividedWires[index] == 'USB 2.0':
            if cir.CircuitPathMode != Electrical.ElectricalCircuitPathMode.AllDevices:  # наикратчайший путь
                cir.CircuitPathMode = Electrical.ElectricalCircuitPathMode.AllDevices
        else:
            # key = 0
            # for e, p in zip(specialElements, specialPanels):
            #     if (e in namesOfTypesOfElements[index]) & (p in namesOfTypesOfPanels[index]):
            #         key = 1
            # print(list(cir.Elements)[0].Id)
            value = doc.GetElement(list(cir.Elements)[0].GetTypeId()).LookupParameter('КабельСиловой').AsString()
            if value:
                if 'пут' in value:
                    if cir.CircuitPathMode != Electrical.ElectricalCircuitPathMode.AllDevices:  # наикратчайший путь
                        cir.CircuitPathMode = Electrical.ElectricalCircuitPathMode.AllDevices
            # else:
            #     if cir.CircuitPathMode != Electrical.ElectricalCircuitPathMode.FarthestDevice:  # путь с диктующей плоскостью
            #         cir.CircuitPathMode = Electrical.ElectricalCircuitPathMode.FarthestDevice
        modes.append(cir.CircuitPathMode)

#------------Длина с запасом и Дискретная длина----------------

lengths = []  # формируем список длин цепей
for cir in (cirs):
    lengths.append(cir.Length)

k = 304.8
cableLengths = [1, 2, 3, 5, 20]  # список возможных длин мерных изделий
cableLengthsForUSB = [3]  # список возможных длин мерных изделий

reserve = []  # Далее формируем список запасов длин цепей
for wire in dividedWires:
    if wire == 'HDMI':
        reserve.append(400)
    elif wire == 'USB 2.0':
        reserve.append(100)
    elif wire == 'VGA':
        reserve.append(100)
    elif wire == 'HDMI(не включать)':
        reserve.append(100)
    elif wire == 'USB 2.0(не включать)':
        reserve.append(100)
    elif wire == '4×Alarm':
        reserve.append(0)
    elif wire == '2×Alarm':
        reserve.append(0)
    else:
        reserve.append(1000)


# Формируем список длин с запасом
lengthsWithReserve = [length * 1.02 + res /
                      k for length, res in zip(lengths, reserve)]

lengthsWithReserve = [ceil(length * k / 1000) * 1000 /
                      k for length in lengthsWithReserve]  # Округление вверх

# округление параметра "длина с запасом" до типовых размеров для кабеля HDMI (и VGA)
for i, length in enumerate(lengthsWithReserve):
    if dividedWires[i] == 'HDMI' or \
       dividedWires[i] == 'HDMI(не включать)' or \
       dividedWires[i] == 'USB 2.0(не включать)' or \
       dividedWires[i] == 'VGA':
        a = 0
        for pos in cableLengths:
            if length <= pos * 1000 / k:
                a = pos * 1000 / k
                break
        if a:
            lengthsWithReserve[i] = a
    else:
        lengthsWithReserve[i] = lengthsWithReserve[i] * 1.5

# Далее формируем список округлённых в большую сторону длин (для мерных изделий)
discreteLengths = []
for i, length in enumerate(lengthsWithReserve):
    if dividedWires[i] == 'HDMI' or \
       dividedWires[i] == 'HDMI(не включать)' or \
       dividedWires[i] == 'USB 2.0(не включать)' or \
       dividedWires[i] == 'VGA':
        # Если длина не будет подобрана, то останутся эти прочерки
        a = 'Error: path length exceeded'
        for pos in cableLengths:
            if length <= pos * 1000 / k:
                a = pos
#                print(111)
                break
        discreteLengths.append(a)
    elif dividedWires[i] == 'USB 2.0':
        # Если длина не будет подобрана, то останутся эти прочерки
        # a = 'Error: path length exceeded'
        a = '3'
        for pos in cableLengthsForUSB:
            if length <= pos * 1000 / k:
                a = pos
#                print(222)
                break
        discreteLengths.append(a)
    else:
        discreteLengths.append('')




for i, cir in enumerate(cirs):
    cir.LookupParameter('КабельЦепи').Set(dividedWires[i])
    cir.LookupParameter('Длина с запасом').Set(lengthsWithReserve[i])
    #cir.LookupParameter('Дискретная длина').Set(str(discreteLengths[i]))


# ---------------Наименование и марка-------------------

nameCode = {  # Словарь строк для "Наименования и технической характеристики"
    'FTP 4x2x0,52 cat 5e': 'Кабель парной скрутки',
    'RG-6U (75 Ом)': 'Кабель коаксильный',
    'UTP 4x2x0,52 cat 5e': 'Кабель парной скрутки',
    'UTP 4x2x0,52 cat 6a': 'Кабель парной скрутки',
    'HDMI': 'Кабель мультимедийный высокоскоростной',
    'HDMI(не включать)': 'Кабель мультимедийный высокоскоростной',
    'USB 2.0': 'Кабель',
    'USB 2.0(не включать)': 'Кабель',
    'VGA': 'Кабель видеоинтерфейсный',
    'Акустический кабель Audiocore Primary Wire M ACS0102': 'Кабель акустический',
    'ВВГнг(А) 3х0,75': 'Кабель силовой',
    'ВВГнг(А)-LS 3x1,5': 'Кабель силовой',
    'SAT 703 B': 'Кабель коаксильный',
    'ШВВП 2х0,75': 'Кабель силовой',
    'КВВГЭ 10х0,75': 'Кабель',
    'FTP 5e': 'Кабель парной скрутки',
    'RG174': 'Кабель коаксильный',
    'RG-174': 'Кабель коаксильный',
    'UTP 4x2x0,52 cat 6e': 'Кабель парной скрутки',
    '3xRCA': 'Кабель передачи аудио-видеосигналов',
    '20 жильный управляющий кабель': 'Кабель (Уточнить)',
    '4×Alarm': 'Кабель специальный',
    '2×Alarm': 'Кабель специальный',
}
markaCode = {  # Словарь строк для "Тип, марка"
    'FTP 4x2x0,52 cat 5e': 'FTP 4x2x0,52 cat 5e',
    'UTP 4x2x0,52 cat 5e': 'UTP 4x2x0,52 cat 5e',
    'UTP 4x2x0,52 cat 6a': 'UTP 4x2x0,52 cat 6a',
    'HDMI': 'HDMI',
    'RG-6U (75 Ом)': 'RG-6U (75 Ом)',
    'HDMI(не включать)': 'HDMI(не включать)',
    'USB 2.0': 'USB 2.0',
    'USB 2.0(не включать)': 'USB 2.0(не включать)',
    'VGA': 'VGA',
    'Акустический кабель Audiocore Primary Wire M ACS0102': 'Audiocore Primary Wire M ACS0102',
    'ВВГнг(А) 3х0,75': 'ВВГнг(А) 3х0,75',
    'ВВГнг(А)-LS 3x1,5': 'ВВГнг(А)-LS 3x1,5',
    'SAT 703 B': 'SAT 703 B',
    'ШВВП 2х0,75': 'ШВВП 2х0,75',
    'КВВГЭ 10х0,75': 'КВВГЭ 10х0,75',
    'FTP 5e': 'FTP 2x2x0,52 cat 5e',
    'RG174': 'RG174',
    'RG-174': 'RG174',
    'UTP 4x2x0,52 cat 6e': 'UTP 4x2x0,52 cat 6e',
    '3xRCA': '3xRCA',
    '20 жильный управляющий кабель': '20 жил (Условно)',
    '4×Alarm': '4×Alarm',
    '2×Alarm': '2×Alarm',
}

naimenovanie = []  # Далее формируем список наименований проводов на основе словаря
for i, wire in enumerate(dividedWires):
    try:
        naimenovanie.append(nameCode[wire])
    except KeyError:
        naimenovanie.append('Ошибка 3: ' + wire)

marka = []  # Далее формируем список марок проводов на основе словаря
for i, wire in enumerate(dividedWires):
    try:
        marka.append(markaCode[wire])
    except KeyError:
        marka.append('Ошибка 4: ' + wire)

marka2 = []  # Далее формируем список марок проводов с учётом мерных изделий
sposobRaschyota = []
# print(len(dividedWires))
# print(len(marka))
# print(len(discreteLengths))
for i, m in enumerate(marka):
    a = discreteLengths[i]
    if a != '':
        marka2.append(m + ', ' + str(a) + ' м')  # например "HDMI, 1.5 м"
        sposobRaschyota.append(1)
    else:
        marka2.append(m)
        sposobRaschyota.append(0)

kolichestvo = []
edinitcyIzmereniia = []
for i, cir in enumerate(cirs):
    if sposobRaschyota[i] == 1:
        kolichestvo.append(1)
        edinitcyIzmereniia.append('шт.')
    else:
        kolichestvo.append(lengthsWithReserve[i] * 304.8 / 1000)
        edinitcyIzmereniia.append('м')


for i, cir in enumerate(cirs):
    cir.LookupParameter('Наименование').Set(naimenovanie[i])
    cir.LookupParameter('Тип, марка').Set(marka2[i])
    cir.LookupParameter('Способ расчёта').Set(sposobRaschyota[i])
    cir.LookupParameter('Количество').Set(kolichestvo[i])
    cir.LookupParameter('Единицы измерения').Set(edinitcyIzmereniia[i])


els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().ToElements()
for el in filter(lambda x: 'Фейк' in x.LookupParameter('Тип').AsValueString(), els):
    doc.Delete(el.Id)
els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().ToElements()
for el in filter(lambda x: 'Фейк' not in x.LookupParameter('Тип').AsValueString(), els):
    el.LookupParameter('Количество').Set(0)
symbols = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsElementType().ToElements()
symbols = list(filter(lambda x: 'Фейк' in x.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString(), symbols))
for symbol in symbols:
    if not symbol.IsActive:
        symbol.Activate()
    symbol.LookupParameter('Описание').Set('')
    symbol.LookupParameter('Комментарии к типоразмеру').Set('')
    symbol.LookupParameter('Ключевая пометка').Set('')
# symbol = list(filter(lambda x: x.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == 'Фейк 3', symbols))[0]
levels = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
level = levels[0]
location = XYZ(-10.8904389851532, 76.6607564383779, 0)
done = []
for i, cir in enumerate(cirs):
    if marka2[i] not in done:
        done.append(marka2[i])
        symbol = symbols[len(done)-1]
        symbol.LookupParameter('Описание').Set(naimenovanie[i])
        symbol.LookupParameter('Комментарии к типоразмеру').Set(marka2[i])
        symbol.LookupParameter('Ключевая пометка').Set(edinitcyIzmereniia[i])
        symbol.LookupParameter('Стоимость').Set(200)
    else:
        symbol = list(filter(lambda x: x.LookupParameter('Комментарии к типоразмеру').AsString() == marka2[i], symbols))[0]
    el = doc.Create.NewFamilyInstance(location, symbol, level, Structure.StructuralType.NonStructural)
    # cir.LookupParameter('Способ расчёта').Set(sposobRaschyota[i])
    el.LookupParameter('Количество').Set(kolichestvo[i])
    el.LookupParameter('Цепь').Set(str(cir.Id))
    el.LookupParameter('Помещение').Set(cir.LookupParameter('Помещение').AsString())
    location += XYZ(0, -0.1, 0)

symbols = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsElementType().ToElements()
symbols = list(filter(lambda x: 'офр' in x.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString(), symbols))

location = XYZ(-10, 76, 0)
els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().ToElements()
for el in filter(lambda x: 'офр' in x.LookupParameter('Тип').AsValueString(), els):
    doc.Delete(el.Id)
els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().ToElements()
for symbol in symbols:
    if not symbol.IsActive:
        symbol.Activate()
    symbol.LookupParameter('Стоимость').Set(300)
    for vega in filter(lambda x: 'Compact A3' in x.LookupParameter('Тип').AsValueString() or 'Блок сетевой управляющий ВЕГА' in x.LookupParameter('Тип').AsValueString(), els):
        el = doc.Create.NewFamilyInstance(location, symbol, level, Structure.StructuralType.NonStructural)
        # print(el)
        kol = symbol.LookupParameter('Группа модели').AsString()
        el.LookupParameter('Количество').Set(float(kol))  # Может быть ошибка, если Группа модели не прописана (её нужно прописывать вручную каждому типу крепежа).
        # el.LookupParameter('Количество').Set(11)
        el.LookupParameter('Помещение').Set(vega.LookupParameter('Помещение').AsString())
        location += XYZ(0, -0.1, 0)





#---------------------------------Линии по трактории-----------------
count = []
circuitPaths = []  # Достаём траектории каждой цепи
for cir in cirs:
    circuitPaths.append(cir.GetCircuitPath())
    count.append(cir.Elements.Size)

pointsAsStr = []  # Перевод данных из незнакомого мне формата в строку
for circuitPath in circuitPaths:
    sublist = []
    for xyz in circuitPath:
        sublist.append(str(xyz).replace('(', '').replace(')', ''))
    pointsAsStr.append(sublist)

k = 304.8   # в 1 футе k мм

points = []  # Из строк делаем Points
for path in pointsAsStr:
    sublist = []
    for i in path:
        x = float(i.split(',')[0])
        y = float(i.split(',')[1])
        z = float(i.split(',')[2])
        # sublist.append(Point().ByCoordinates(x * k, y * k, z * k))
        sublist.append(XYZ(x, y, z))
    points.append(sublist)

lines = []  # Из Points делаем Lines
for path, c in zip(points, count):
    # if c > 1:  # Лучевая схема
    if 2 > 1:  # Ортогональная схема
        sublist = []
        for i, p in enumerate(path):
            try:
                start, end = path[i], path[i + 1]

                line = Line.CreateBound(start, end)  #https://forums.autodesk.com/t5/revit-api-forum/how-to-crete-3d-modelcurves-to-avoid-exception-curve-must-be-in/td-p/8355936
                direction = line.Direction
                x, y, z = direction.X, direction.Y, direction.Z
                normal = XYZ(z - y, x - z, y - x)
                # normal = XYZ.BasisZ.CrossProduct(line.Direction)
                plane = Plane.CreateByNormalAndOrigin(normal, start)
                sketchPlane = SketchPlane.Create(doc, plane)
                modelCurve = doc.Create.NewModelCurve(line, sketchPlane)

                # sublist.append(Line().ByStartPointEndPoint(path[i], path[i + 1]))
                sublist.append(modelCurve)
            except IndexError:
                continue
        lines.append(sublist)
    else:
        sublist = []
        sublist.append(Line.ByStartPointEndPoint(path[0], path[-1]))
        lines.append(sublist)

OUT = cirs, panels, elements, dividedWires, lengthsWithReserve, discreteLengths, naimenovanie, marka2, sposobRaschyota, kolichestvo, edinitcyIzmereniia,


electricalEquipment = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().ToElements()

for el in electricalEquipment:
    # print(el.Id)
    if el.GetSubComponentIds():
        room_name = el.LookupParameter('Помещение').AsString()
        for sub_el_id in el.GetSubComponentIds():
#            print(doc.GetElement(sub_el_id).LookupParameter('Тип').AsValueString())
            doc.GetElement(sub_el_id).LookupParameter('Помещение').Set(room_name)














import Autodesk

paths = lines

categories = doc.Settings.Categories # получаем все категории
lineCat = categories.get_Item(BuiltInCategory.OST_Lines )  # из всех категорий выбираем категорию линий (о чем свидетельствует OST_Lines)
lineStyleSubTypes = lineCat.SubCategories # Все стили линий лежат в так называемой субкатегории, можно убедиться и вывести их имена OUT =  [i.Name for i in lineStyleSubTypes]

lineStyles = [i for i in FilteredElementCollector(doc).OfClass(Autodesk.Revit.DB.GraphicsStyle)]

"""nameCode = [
    'UTP 2x2x0,52 cat 5e',
    'FTP 2x2x0,52 cat 5e',
    'HDMI',
    'USB 2.0',
    'VGA',
    'Акустический кабель Audiocore Primary Wire M ACS0102',
    'ВВГнг(А) 3х0,75',
    'ВВГнг(А)-LS 3x1,5',
    'Коакс. кабель (SDI)',
    'SAT 703 B',
    'ШВВП 2х0,75',
    'КВВГЭ 10х0,75',
    'FTP 5e',
    'RG174',
    'RG-174',
    '20 жильный управляющий кабель']
"""

IN1 = dividedWires

nameCode = []
for i in IN1:
    if i not in nameCode: nameCode.append(i)
nameCode.sort()

# Создадим новые стили:

for name in nameCode:
    try:
        newLineStyleCat = categories.NewSubcategory(lineCat, '!мвс ' + name )
    except:
        continue


myStyles = []
myStylesNames = []

for style in lineStyles:
    if style.Name.ToString().replace('!мвс ','') in nameCode:
        myStyles.append(style)
        myStylesNames.append(style.Name)

try:
    errCounter = []
    list1 = []
    for i, path in enumerate(paths):
        sublist = []
        for line in path:
            try:
                line.LineStyle = myStyles[nameCode.index(IN1[i])]
                sublist.append(line.LineStyle.Name)
            except ValueError:
                s = str(i) + ' ' + IN1[i]
                if s not in errCounter: errCounter.append(s)
                sublist.append('Кабель не найден: ' + IN1[i])
        list1.append(sublist)
except:
    pass

#OUT2 = errCounter, list1, sorted(nameCode), sorted(myStylesNames)



els = FilteredElementCollector(doc).OfCategory(
    BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().ToElements()


mydict = {}
for el in els:
    symbol = doc.GetElement(el.GetTypeId())
    if symbol.LookupParameter('Изображение типоразмера').AsElementId().ToString() == '-1':
        if symbol.LookupParameter('URL').AsString() != 'Без УГО':
            name = symbol.LookupParameter('Имя типа').AsString()
            if name not in mydict.keys():
                mydict[name] = []
            mydict[name].append(el.Id)

# print(mydict)

if mydict:
    if len(mydict.keys()) == 1:
        print('Данный типоразмер не имеет УГО. Следует добавить Изображение типоразмера (310×167.png) либо в параметре URL указать "Без УГО"')
    else:
        print('Данные типоразмеры не имеют УГО. Следует добавить Изображение типоразмера (310×167.png) либо в параметре URL указать "Без УГО"')
    for name in mydict.keys():
        print('{}'.format(output.linkify(mydict[name], name)))





t.Commit()
