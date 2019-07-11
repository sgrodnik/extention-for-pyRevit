# -*- coding: utf-8 -*-
""""""
__title__ = 'Схема'
__author__ = 'SG'

import os
import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *
from math import ceil

from Autodesk.Revit.DB import XYZ, FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

k = 304.8
# print(doc.PathName)
path = '\\'.join(doc.PathName.split('\\')[:-1]) + '\\' + doc.PathName.split('\\')[-1].split('.')[0] + '.txt'
try:
    src = open(path, 'r').read().decode("utf-8")[:-1]
    # src2 = open(path, 'r').read().decode("UTF-8")[:-1]
    # src2 = open(path, 'r').read()[:-1]
    # print(src)
except IOError:
    info = path + '\nЭтот файл должен содержать таблицу (разделители - табуляция).\nПервая строка должна содержать заголовки.\nПоследняя строка должна быть пустой.\n'
    # file = open(path, 'w')
    with open(path, 'w') as file:
        file.write(info.encode("utf-8"))
    # file.close()
    os.startfile(path)

# # src = open(path, 'r', encoding="utf-8").read()[:-1]

# print('-----------------------------------------------------------')
# print(src2)

titles = src.split('\n')[0].split('\t')
_area = titles.index('Площадь, м²')
_number = titles.index('№')
_supplyAirFlow = titles.index('Расход приточного воздуха, м³/ч')
_supplySystem = titles.index('Приточная система')
_exhaustSystem = titles.index('Вытяжная система')
_exhaustAirFlow = titles.index('Расход удаляемого воздуха, м³/ч')
_exhaustDevices = titles.index(' Вытяжные  воздухораспределительные устройства')
_supplyDevices = titles.index(' Приточные  воздухораспределительные устройства')
_sortFactor = titles.index('Сортировка')

# devs = {'': 000,
#         }


# def countDevices(room):
#     roomFlow = room.supplyAirFlow
#     unitFlow = devs[room.supplyDevices]
#     number = roomFlow / unitFlow
#     return number


class Room:
    allRefsByGroup = {}

    def __init__(self, string=' \t' * 15):
        raw = string.split('\t')
        if raw[_number] != ' ':
            nmbr = raw[_number]
            if raw[_sortFactor] not in Room.allRefsByGroup.keys():
                Room.allRefsByGroup[raw[_sortFactor]] = set()
            Room.allRefsByGroup[raw[_sortFactor]].add(self)
        else:
            nmbr = 'Empty'
            print(nmbr + ' did not add')
        self.number = nmbr
        self.area = raw[_area]
        self.supplyAirFlow = raw[_supplyAirFlow] if raw[_supplyAirFlow] else '0'
        self.exhaustAirFlow = raw[_exhaustAirFlow] if raw[_exhaustAirFlow] else '0'
        self.supplyDevices = raw[_supplyDevices]
        self.exhaustDevices = raw[_exhaustDevices] if 'Выт' not in raw[_exhaustDevices] else 'Шкаф'
        self.supplySystem = raw[_supplySystem] if raw[_supplySystem] else 'ЯЯЯ'
        self.exhaustSystem = raw[_exhaustSystem].replace('.', 'ЯЯЯ')

    def __str__(self):
        return self.number


for raw in src.split('\n')[1:]:
    try:
        group = Room.allRefsByGroup[raw[_sortFactor]]
    except KeyError:
        group = []
    if raw.split('\t')[_number] not in map(lambda x: x.number, group):
        Room(raw)
    else:
        raw = raw.split('\t')
        room = filter(lambda x: x.number == raw[_number], Room.allRefsByGroup[raw[_sortFactor]])[0]
        room.supplyAirFlow += ', ' + raw[_supplyAirFlow] if raw[_supplyAirFlow] else ''
        room.exhaustAirFlow += ', ' + raw[_exhaustAirFlow] if raw[_exhaustAirFlow] else ''
        room.supplyDevices += ', ' + raw[_supplyDevices] if raw[_supplyDevices] else ''
        room.exhaustDevices += ', ' + raw[_exhaustDevices] if 'Выт' not in raw[_exhaustDevices] else 'Шкаф'
        room.supplySystem += ', ' + raw[_supplySystem] if raw[_supplySystem] else ''
        room.exhaustSystem += ', ' + raw[_exhaustSystem] if raw[_exhaustSystem] else ''

pFamilys = {'FFD2024': 'Ламинар',
            'FFD2424': 'Ламинар',
            'FFD3030': 'Ламинар',
            'H11': 'H11',
            'H13': 'H13',
            'H14': 'H14',
            'THLabor2416': 'Ламинар',
            'ДПУ приток': 'ДПУ приток',
            'ДПУ': 'ДПУ',
            '4АПН+ЗКСД': '4АПН',
            'Шкаф': 'Шкаф',
            # '': '',
            }

t = Transaction(doc, 'Схема')
t.Start()

activeView = doc.ActiveView
anns = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericAnnotation).WhereElementIsElementType().ToElements()
roomFamily = list(filter(lambda x: x.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == 'Помещение', anns))[0]

yCoord = 0
for groupName in sorted(Room.allRefsByGroup.keys()):
    lst = sorted(Room.allRefsByGroup[groupName], key=lambda x: x.number)
    lst = sorted(lst, key=lambda x: x.exhaustSystem)
    lst = sorted(lst, key=lambda x: x.supplySystem)

    xCoord = 0
    print('----------------------------------------------------------------')
    print(groupName)
    for room in lst:
        roomInstance = doc.Create.NewFamilyInstance(XYZ(xCoord / k, yCoord / k, 0), roomFamily, activeView)
        info = ''
        for attrName in room.__dict__:
            info += getattr(room, attrName) + '\n'
        print(info)
        string = room.number
        roomInstance.LookupParameter('Номер').Set(string.replace('ЯЯЯ', '').replace('\n0\n', '\n\n'))
        roomInstance.LookupParameter('Высота').Set(20 / k)
        roomInstance.LookupParameter('Приток').Set(room.supplySystem.replace('ЯЯЯ', ''))
        roomInstance.LookupParameter('Вытяжка').Set(room.exhaustSystem.replace('ЯЯЯ', ''))
        # el.LookupParameter('Номер').Set('{}\n{}\n{}'.format(i.number, i.supplySystem, i.exhaustSystem))
        supplyDevices = 'ДПУ приток' if room.supplyDevices == 'ДПУ' and int(room.supplyAirFlow) else room.supplyDevices
        if supplyDevices:
            pFamilyName = pFamilys[supplyDevices]
            pFamily = list(filter(lambda x: x.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == pFamilyName, anns))[0]
            deviceInstance = doc.Create.NewFamilyInstance(XYZ((xCoord + 3) / k, (yCoord + 20) / k, 0), pFamily, activeView)
        exhaustDevices = room.exhaustDevices
        i = 0
        for exhaustDevice in exhaustDevices.split(', '):
            if exhaustDevice:
                pFamilyName = pFamilys[exhaustDevice]
                pFamily = list(filter(lambda x: x.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == pFamilyName, anns))[0]
                deviceInstance = doc.Create.NewFamilyInstance(XYZ((xCoord + 9 + i) / k, (yCoord + 20) / k, 0), pFamily, activeView)
                i += 6
        width = 12
        width = width if i <= 6 else width + i - 6
        roomInstance.LookupParameter('Ширина').Set(width / k)
        xCoord += width
    yCoord -= 70

t.Commit()
