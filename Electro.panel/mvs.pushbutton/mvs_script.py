# -*- coding: utf-8 -*-
"""(Добавить описание и инструкцию)"""
__title__ = 'Расчет\nMVS'
__author__ = 'SG'

import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *
from Autodesk.Revit.DB import ElementId, PartUtils, ViewOrientation3D, XYZ, FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, Line, Point, SketchPlane, Plane
import sys
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
# import math
from math import ceil

cirs = FilteredElementCollector(doc).OfCategory(
    BuiltInCategory.OST_ElectricalCircuit).WhereElementIsNotElementType().ToElements()

# if 1: #
#   import clr
#   clr.AddReference("RevitServices")

#   import RevitServices
#   from RevitServices.Persistence import DocumentManager
#   from RevitServices.Transactions import TransactionManager

#   doc = DocumentManager.Instance.CurrentDBDocument

#   clr.AddReference('RevitAPI')
#   import Autodesk
#   from Autodesk.Revit.DB import *

#   clr.AddReference('ProtoGeometry')
#   from Autodesk.DesignScript.Geometry import *


firstElementsOfCirs = [list(cir.Elements)[0].Id for cir in cirs]

pairs = [(key, cir) for key, cir in zip(firstElementsOfCirs, cirs)]

pairs.sort(key=lambda x: x[0])

sortedCirs = [p[1] for p in pairs]

cirs = [i for i in sortedCirs]

panels = []  # Далее формируем список панелей цепей
elements = []  # Далее формируем двумерный список элементов цепей
for cir in (cirs):
    panels.append(cir.BaseEquipment)  # Возвращает элемент
    elements.append(cir.Elements)  # Возвращает список(!) элементов

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
            BuiltInParameter.SYMBOL_NAME_PARAM).AsString())
        if not el.LookupParameter('Помещение').AsString():
            raise Exception('Следует заполнить параметр Помещение')
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
"""
numberOfWires = [] # Позднее нам потребуется знать количество разделителей
for wire in wires:
    if wire.find('/') == -1:
        numberOfWires.append(0)       --------УДАЛИТЬ ЭТОТ КУСОК-----------
    else:
        numberOfWires.append(wire.count('/'))"""

doneList = []
dividedWires = [i for i in wires]
# Хитрый цикл, в котором стринги кабелей с разделителями заменяем отдельными кабелями
for i, wire in enumerate(dividedWires):
    if wire.count('/') > 0:
        if i not in doneList:
            for j, w in enumerate(wire.split('/')):
                try:
                    dividedWires.pop(i + j)
                except:
                    a = 'похер'
                dividedWires.insert(i + j, w)
                doneList.append(i + j)

#-------------Смена режима траектории--------------------

specialElements = ['Рабочая станция', 'Конвертер 3G-SDI/HDMI', 'Розетка HDMI', 'Монитор управления',
                   'Консоль (розетка HDMI)', 'Переходник DVI-D/HDMI', 'Розетка HDMI', 'Масштабатор', 'Монитор управления', 'USB2']  # список оборудования, от которого прокладка ведётся по кратчайшему пути
specialPanels = ['MiniPC', 'Монитор обзорный', 'Конвертер HDMI/SDI', 'MiniPC', 'Конвертер HDMI/SDI', 'Розетка HDMI',
                 'Масштабатор', 'Конвертер HDMI/SDI', 'USB1', 'Compact A3']  # список панелей, к которым прокладка ведётся по кратчайшему пути

IN = 0
if IN:
    modes = []
    for index, cir in enumerate(cirs):
        if dividedWires[index] == 'USB 2.0':
            if cir.CircuitPathMode != Autodesk.Revit.DB.Electrical.ElectricalCircuitPathMode.AllDevices:  # наикратчайший путь
                cir.CircuitPathMode = Autodesk.Revit.DB.Electrical.ElectricalCircuitPathMode.AllDevices
        else:
            key = 0
            for e, p in zip(specialElements, specialPanels):
                if (e in namesOfTypesOfElements[index]) & (p in namesOfTypesOfPanels[index]):
                    key = 1
            if key:
                if cir.CircuitPathMode != Autodesk.Revit.DB.Electrical.ElectricalCircuitPathMode.AllDevices:  # наикратчайший путь
                    cir.CircuitPathMode = Autodesk.Revit.DB.Electrical.ElectricalCircuitPathMode.AllDevices
            else:
                if cir.CircuitPathMode != Autodesk.Revit.DB.Electrical.ElectricalCircuitPathMode.FarthestDevice:  # путь с диктующей плоскостью
                    cir.CircuitPathMode = Autodesk.Revit.DB.Electrical.ElectricalCircuitPathMode.FarthestDevice
        modes.append(cir.CircuitPathMode)

#------------Длина с запасом и Дискретная длина----------------

lengths = []  # формируем список длин цепей
for cir in (cirs):
    lengths.append(cir.Length)

k = 304.8
cableLengths = [1, 20]  # список возможных длин мерных изделий
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
    else:
        reserve.append(2000)

# Формируем список длин с запасом
lengthsWithReserve = [len * 1.02 + res /
                      k for len, res in zip(lengths, reserve)]

lengthsWithReserve = [ceil(len * k / 1000) * 1000 /
                      k for len in lengthsWithReserve]  # Округление вверх

# округление параметра "длина с запасом" до типовых размеров для кабеля HDMI (и VGA)
for i, len in enumerate(lengthsWithReserve):
    if dividedWires[i] == 'HDMI' or \
       dividedWires[i] == 'HDMI(не включать)' or \
       dividedWires[i] == 'USB 2.0(не включать)' or \
       dividedWires[i] == 'VGA':
        a = 0
        for pos in cableLengths:
            if len <= pos * 1000 / k:
                a = pos * 1000 / k
                break
        if a:
            lengthsWithReserve[i] = a
    else:
        lengthsWithReserve[i] = lengthsWithReserve[i] * 1.5

# Далее формируем список округлённых в большую сторону длин (для мерных изделий)
discreteLengths = []
for i, len in enumerate(lengthsWithReserve):
    if dividedWires[i] == 'HDMI' or \
       dividedWires[i] == 'HDMI(не включать)' or \
       dividedWires[i] == 'USB 2.0(не включать)' or \
       dividedWires[i] == 'VGA':
        # Если длина не будет подобрана, то останутся эти прочерки
        a = 'Error: path length exceeded'
        for pos in cableLengths:
            if len <= pos * 1000 / k:
                a = pos
                break
        discreteLengths.append(a)
    elif dividedWires[i] == 'USB 2.0':
        # Если длина не будет подобрана, то останутся эти прочерки
        a = 'Error: path length exceeded'
        for pos in cableLengthsForUSB:
            if len <= pos * 1000 / k:
                a = pos
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
    '20 жильный управляющий кабель': 'Кабель (Уточнить)'
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
    '20 жильный управляющий кабель': '20 жил (Условно)'
}

naimenovanie = []  # Далее формируем список наименований проводов на основе словаря
for i, wire in enumerate(dividedWires):
    try:
        naimenovanie.append(nameCode[wire])
    except KeyError:
        naimenovanie.append('Ошибка 3')

marka = []  # Далее формируем список марок проводов на основе словаря
for i, wire in enumerate(dividedWires):
    try:
        marka.append(markaCode[wire])
    except KeyError:
        marka.append('Ошибка 4')

marka2 = []  # Далее формируем список марок проводов с учётом мерных изделий
sposobRaschyota = []
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

# print(lines)




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


OUT2 = errCounter, list1, sorted(nameCode), sorted(myStylesNames)












t.Commit()
