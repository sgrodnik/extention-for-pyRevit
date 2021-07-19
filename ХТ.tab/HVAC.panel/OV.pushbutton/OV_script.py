# -*- coding: utf-8 -*-
""""""
__title__ = 'Расчёт\nспеки'
__author__ = 'SG'


from pyrevit import script
import math
import time
startTime = time.time()
import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *

from Autodesk.Revit.DB import Mechanical, FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

t = Transaction(doc, 'Настройка префиксов и суффиксов воздуховодов')
t.Start()

settings = Mechanical.DuctSettings.GetDuctSettings(doc)
if settings.RoundDuctSizePrefix != 'ø':
    settings.RoundDuctSizePrefix = 'ø'
if settings.RoundDuctSizeSuffix != '':
    settings.RoundDuctSizeSuffix = ''
if settings.RectangularDuctSizeSeparator != '×':
    settings.RectangularDuctSizeSeparator = '×'

t.Commit()

t = Transaction(doc, 'Расчёт спеки')
t.Start()

ducts = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctCurves).WhereElementIsNotElementType().ToElements()  # Воздуховоды
dFits = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctFitting).WhereElementIsNotElementType().ToElements()  # Соединительные детали воздуховодов
terms = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctTerminal).WhereElementIsNotElementType().ToElements()  # Воздухораспределители
flexes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FlexDuctCurves).WhereElementIsNotElementType().ToElements()  # Гибкие воздуховоды
dArms = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctAccessory).WhereElementIsNotElementType().ToElements()  # Арматура воздуховодов
insuls = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctInsulations).WhereElementIsNotElementType().ToElements()  # Материалы изоляции воздуховодов
pipes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType().ToElements()  # Трубы
pFits = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeFitting).WhereElementIsNotElementType().ToElements()  # Соединительные детали трубопроводов
pArms = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).WhereElementIsNotElementType().ToElements()  # Арматура трубопроводов
pipeInsuls = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeInsulations).WhereElementIsNotElementType().ToElements()  # Материалы изоляции труб
equipments = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType().ToElements()  # Оборудование
fakes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()  # Обобщенные модели
piping_systems = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipingSystem).WhereElementIsNotElementType().ToElements()  # Обобщенные модели
sanitary = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PlumbingFixtures).WhereElementIsNotElementType().ToElements()  # Оборудование
[equipments.Add(i) for i in sanitary]

k = 304.8
rad = 57.2957795130823

rects = []  # Воздуховоды прям
rounds = []  # Воздуховоды кругл
for i in ducts:
    if i.LookupParameter('Ширина'):
        rects.append(i)
    elif i.LookupParameter('Диаметр'):
        rounds.append(i)

rectsSize = []  # Размер воздуховодов
for i in rects:
    b = round(i.LookupParameter('Ширина').AsDouble() * k, 2)
    h = round(i.LookupParameter('Высота').AsDouble() * k, 2)
    if h > b:
        b, h = h, b
    if b < 251:
        s = 'δ=0,5'
    elif b < 1001:
        s = 'δ=0,7'
    else:
        s = 'δ=0,9'
    sys_name = i.LookupParameter('Имя системы').AsString()
    sys_name = sys_name if sys_name else 'Не определено'
    if 'Д' in sys_name:
        s = 'δ=1,2'
    if i.LookupParameter('Длина').AsDouble() * k > 10:
        rectsSize.append('{:.0f}×{:.0f}, {}'.format(b, h, s))
    else:
        rectsSize.append('Не учитывать {:.0f}×{:.0f}, {}'.format(b, h, s))

roundsDiameter = []  # Диаметр воздуховодов (и гибких)
for i in rounds:
    d = round(i.LookupParameter('Диаметр').AsDouble() * k, 2)
    if d < 201:
        s = 'δ=0,5'
    elif d < 451:
        s = 'δ=0,6'
    else:
        s = 'δ=0,7'
    if i.LookupParameter('Имя системы').AsString():
        if 'Д' in i.LookupParameter('Имя системы').AsString():
            s = 'δ=1,2'
    else:
        s = ''
    roundsDiameter.append('ø{:.0f}, {}'.format(d, s))

els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FlexDuctCurves).WhereElementIsElementType().ToElements()
symbol = list(filter(lambda x: x.LookupParameter('Имя типа').AsString() == 'SG_Круглый изолированный', els))[0]
flexDiameter = []
for flex in flexes:
    flexDiameter.append('ø{:.0f}'.format(flex.LookupParameter('Диаметр').AsDouble() * k))

    if flex.LookupParameter('Классификация систем').AsString() == 'Приточный воздух' \
        and flex.LookupParameter('Тип').AsValueString() != 'SG_Круглый изолированный':
            flex.FlexDuctType = symbol

lengthDucts = []  # Длина воздуховодов
for i in ducts:
    lengthDucts.append(round(i.LookupParameter('Длина').AsDouble() * 1.1, 2))
lengthFlexes = []
for i in flexes:
    lengthFlexes.append(i.LookupParameter('Длина').AsDouble() * 1.15 + 100 / k)
lengthPipes = []
for i in pipes:
    lengthPipes.append(round(i.LookupParameter('Длина').AsDouble() * 1.1, 2))

kk = 10.76391041671  # Площадь изоляции воздуховодов и фитингов
kkk = k * kk / 1000
hosts = []
hostName = []
hostsCat = []
hostsArea = []
insArea = []
insThickness = []
for i in insuls:
    host = doc.GetElement(i.HostElementId)
    hosts.append(host)
    hostName.append(host.Name)
    hostsCat.append(host.get_Parameter(BuiltInParameter.ELEM_CATEGORY_PARAM).AsValueString())
    if host.LookupParameter('ХТ Площадь'):
        hostsArea.append(host.LookupParameter('ХТ Площадь').AsDouble() / kkk * 1.1)
    else:
        hostsArea.append('')
    if i.LookupParameter('Площадь'):
        if i.LookupParameter('Площадь').AsDouble() > 0:
            insArea.append(i.LookupParameter('Площадь').AsDouble() / kkk * 1.1)
        else:
            insArea.append('non')
    else:
        insArea.append('-')
    # if str(i.Id) == '741520':
    #     print(i.LookupParameter('Площадь').AsValueString())
    #     print(i.LookupParameter('Площадь').AsDouble() / kkk * 1.1 * k)

    insType = i.Name
    if host.LookupParameter('Комментарии').AsString() == 'Утепление+':
        insThickness.append('δ=32')
#   insType = doc.GetElement(i.GetTypeId()).get_Parameter(BuiltInParameter.ELEM_CATEGORY_PARAM).AsValueString()
    elif 'пожар' in insType or 'гнеза' in insType:
        # insThickness.append('IE30')
        koms = doc.GetElement(i.GetTypeId()).LookupParameter('Комментарии к типоразмеру').AsString()
        if koms:
            insThickness.append(koms)
        else:
            insThickness.append('')
        # insThickness.append('IE60, δ=20') # Для медси (для других надо удалить). Надо продумать аналогию с оборудой
    else:
        insThickness.append('δ=10')

areaInsuls = []
for i in range(len(insuls)):
    if hostsArea[i] != '':
        areaInsuls.append(hostsArea[i])
    elif insArea[i] != 'non':
        areaInsuls.append(insArea[i])
    else:
        areaInsuls.append(0)


# with open('C:\\Users\\SG\\Desktop\\log.log', 'w') as f:

sizes = []  # Соединители и арматура
for i in list(dFits) + list(dArms):
    size = i.LookupParameter('Размер').AsString()
    size.replace(' мм', '')
    if size[:(size.find("-"))] == size[(size.find("-") + 1):]:
        size = size[:(size.find("-"))]

    # if size[-1] == 'ø':
    #   size = 'ø' + size[:-1]

    if i.LookupParameter('Смещение S'):
        offset = i.LookupParameter('Смещение S').AsDouble()
        offset = offset * k
        size = '{}, S={:.0f}'.format(size, offset)

    if i.LookupParameter('Угол отвода'):
        angle = i.LookupParameter('Угол отвода').AsDouble() * rad
        if angle > 0.5:
            size = '{}, {:.0f}°'.format(size, angle)
#   typeMarka = doc.GetElement(i.GetTypeId()).LookupParameter('Тип, марка, обозначение документа, опросного листа').AsString() Тут место для Маркировка типоразмера
#   if typeMarka: size = typeMarka

#  -------------------------------------------------------------- Обработать круглые врезки

    name = doc.GetElement(i.GetTypeId()).get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    if 'ФБО' in name:
        size = name
    elif 'НП' in name:
        size = name
    # f.write('{}\n'.format(i.Id))
    opis = doc.GetElement(i.GetTypeId()).LookupParameter('Описание').AsString()
    if opis:
        if 'Переход стальной с прямоугольного на прямоугольное сечение' in opis:
            x1 = round(i.LookupParameter('Ширина воздуховода 1').AsDouble() * k, 2)
            x2 = round(i.LookupParameter('Ширина воздуховода 2').AsDouble() * k, 2)
            y1 = round(i.LookupParameter('Высота воздуховода 1').AsDouble() * k, 2)
            y2 = round(i.LookupParameter('Высота воздуховода 2').AsDouble() * k, 2)
            dx = round(i.LookupParameter('ШиринаСмещения').AsDouble() * k, 2)
            dy = round(i.LookupParameter('ВысотаСмещения').AsDouble() * k, 2)

            X1 = x1
            X2 = x2
            Y1 = y1
            Y2 = y2
            DX = dx
            DY = dy
            f1 = X1 * Y1
            f2 = X2 * Y2
            width = X1 >= X2
            maxf = f1 >= f2
            height = Y1 >= Y2
            total = width + maxf + height
            revX = 1
            if total == 1:
                revX = -1
            if total == 2:
                revX = -1
            revY = -1
            if f1 >= f2:
                revY = 1

            if abs(DX - X1 / 2) < 2:
                xxx = 'center'  # Допустимая погрешность 2 мм
            elif abs(DX - X1 / 2 + abs(X1 - X2) / 2 * revX) < 2:
                xxx = 'left'
            elif abs(DX - X1 / 2 - abs(X1 - X2) / 2 * revX) < 2:
                xxx = 'right'
            else:
                xxx = 'error'

            if abs(DY - Y1 / 2) < 2:
                yyy = 'center'
            elif abs(DY - Y1 / 2 + abs(Y1 - Y2) / 2 * revY) < 2:
                yyy = 'up'
            elif abs(DY - Y1 / 2 - abs(Y1 - Y2) / 2 * revY) < 2:
                yyy = 'down'
            else:
                yyy = 'error'

            x, y = xxx, yyy

            if y == 'up':
                if x == 'left': type_ = 'Тип 1'
                if x == 'center': type_ = 'Тип 2'
                if x == 'right': type_ = 'Тип 3'
                if x == 'error': type_ = 'Error'
            if y == 'down':
                if x == 'left': type_ = 'Тип 3'
                if x == 'center': type_ = 'Тип 2'
                if x == 'right': type_ = 'Тип 1'
                if x == 'error': type_ = 'Error'
            if y == 'center':
                if x == 'left': type_ = 'Тип 2 поворот'
                if x == 'center': type_ = 'Тип 4'
                if x == 'right': type_ = 'Тип 2 поворот'
                if x == 'error': type_ = 'Error'
            if y == 'error': type_ = 'Error'

            sizeAsList = size.replace('-', '×').split('×')

            try:
                x1 = sizeAsList[0]
                y1 = sizeAsList[1]
                x2 = sizeAsList[2]
                y2 = sizeAsList[3]
            except:
                type_ = 'Error:Equal'
                x2 = '0'
                y2 = '0'

            if 'поворот' in type_:
                size = y1 + '×' + x1 + '-' + y2 + '×' + x2 + ', Тип 2'
            else:
                size = x1 + '×' + y1 + '-' + x2 + '×' + y2 + ', ' + type_

    sizes.append(size)

systemNamesVent = []
for i in list(dFits) + list(terms) + list(dArms) + list(ducts) + list(flexes) + list(insuls):
    if i.LookupParameter('Имя системы').AsString():
        systemNamesVent.append(i.LookupParameter('Имя системы').AsString().split('/')[0])
    else:
        systemNamesVent.append('Не определено')

donelist = []
for i in terms:
    if i.LookupParameter('Семейство').AsValueString() == 'Шкаф вытяжной с ребрами':
        h = i.LookupParameter('Смещение').AsDouble()
        try:
            i.LookupParameter('Высота потолка').Set(h if h * k > 2000 else 2400 / k)
        except AttributeError:
            pass
        symbol = doc.GetElement(i.GetTypeId())
        # 300×100; 250×100h; 250×150h
        if symbol.Id not in donelist:
            height_bottom = int(symbol.LookupParameter('Высота нижнего отверстия').AsDouble() * k)
            height_bottom = height_bottom if height_bottom > 100 else 100
            height_top = int(symbol.LookupParameter('Высота верхнего отверстия').AsDouble() * k)
            height_top = height_top if height_top > 100 else 100
            name = str(int(symbol.LookupParameter('Ширина шкафа').AsDouble() * k))
            name += '×' + str(int(symbol.LookupParameter('Глубина шкафа').AsDouble() * k))
            name += '; Отверстия: ▲' + str(int(symbol.LookupParameter('Ширина верхнего отверстия').AsDouble() * k))
            name += '×' + str(height_top)
            name += 'h, ▼' + str(int(symbol.LookupParameter('Ширина нижнего отверстия').AsDouble() * k))
            name += '×' + str(height_bottom)
            name += 'h;'
            name += '\nполная высота стены {:.0f} мм'.format(i.LookupParameter('Смещение').AsDouble() * k)
            symbol.LookupParameter('Группа модели').Set(name)
            donelist.append(symbol.Id)


sizesOfTerms = []
for i in terms:
    mark = ''
    typeMarka = doc.GetElement(i.GetTypeId()).LookupParameter('Группа модели').AsString()
    if typeMarka:
        mark = typeMarka
    sizesOfTerms.append(mark)


    # -------------------------------------- Трубы --------------------------------------------
    # -------------------------------------- Трубы --------------------------------------------
    # -------------------------------------- Трубы --------------------------------------------
    # -------------------------------------- Трубы --------------------------------------------
    # -------------------------------------- Трубы --------------------------------------------
    # -------------------------------------- Трубы --------------------------------------------
    # -------------------------------------- Трубы --------------------------------------------
    # -------------------------------------- Трубы --------------------------------------------
    # -------------------------------------- Трубы --------------------------------------------
    # -------------------------------------- Трубы --------------------------------------------
    # -------------------------------------- Трубы --------------------------------------------
    # -------------------------------------- Трубы --------------------------------------------
    # -------------------------------------- Трубы --------------------------------------------
    # -------------------------------------- Трубы --------------------------------------------
    # -------------------------------------- Трубы --------------------------------------------
    # -------------------------------------- Трубы --------------------------------------------


# pipesHolod = []  # -------------------------------------- Трубы -------------------------
# pipesOther = []
# for i in pipes:
#     if i.LookupParameter('Имя системы').AsString():
#         system = i.LookupParameter('Имя системы').AsString().split('/')[0]
#         if 'Х' in system and 'Х1.' not in system and 'Х2.' not in system and 'Х3.' not in system and 'Х4.' not in system:  # Кириллическая Ха
#             pipesHolod.append(i)
#         else:
#             pipesOther.append(i)

pipesDict = {
    6.35 : '1/4" (6,35 mm)',
    9.52 : '3/8" (9,52 mm)',
    12.7 : '1/2" (12,7 mm)',
    15.88: '5/8" (15,9 mm)',
    15.9 : '5/8" (15,9 mm)',
    19.05: '3/4" (19,05 mm)',
    22.23: '7/8" (22,23 mm)',
    25.4 : '1" (25,4 mm)',
    28.58: '1-1/8" (28,58 mm)',
    31.75: '1-1/4" (31,75 mm)',
    34.93: '1-3/8" (34,93 mm)',
    41.28: '1-5/8" (41,28 mm)'
}

# systemNamesOther = []
# sizesOfPipesOther = []
# for i in pipesOther:
#     system = i.LookupParameter('Имя системы').AsString().split('/')[0].replace('обр', '').replace('под', '').replace('Xо ', 'X').replace('Xп ', 'X').replace('Xд ', 'X')
#     system = system.replace('Х3.', 'Х.').replace('Х4.', 'Х.').replace('Хд.', 'Х.')
#     systemNamesOther.append(system)
#     size = i.LookupParameter('Диаметр').AsDouble() * k
#     if 'T1 ' in system or 'T2 ' in system:  # Латинская t
#         if system.split()[1] in systemOtherDict:
#             sizesOfPipesOther.append(systemOtherDict[system.split()[1]])
#         else:
#             if 'T1 К' in system or 'T2 К' in system:
#                 sizesOfPipesOther.append('ø{:.0f}'.format(size))
#             else:
#                 sizesOfPipesOther.append('Не определено')
#     elif size in pipesDict:
#         sizesOfPipesOther.append(pipesDict[size])
#     else:
#         sizesOfPipesOther.append('ø{:.0f}'.format(size))

# systemDict = {
#     '101': 'ø40',
#     '102': 'ø50',
#     '103': 'ø50',
#     '106': 'ø50',
#     '107': 'ø50',
#     '107а': 'ø40',
#     '108': 'ø40'
# }

# systemNamesHolod = []
# sizesOfPipesHolod = []
# for i in pipesHolod:
#     systemName = i.LookupParameter('Имя системы').AsString().split('/')[0].replace('обр', '').replace('под', '').replace('Xо ', 'X').replace('Xп ', 'X').replace('Xд ', 'X')
#     systemName = systemName.replace('Х3.', 'Х.').replace('Х4.', 'Х.').replace('Хд.', 'Х.')
#     systemNamesHolod.append(systemName)
#     if 'Х1' == systemName or 'Х2' == systemName:  # Системы магистальных участков
#         size = i.LookupParameter('Диаметр').AsDouble() * k
#         sizesOfPipesHolod.append('ø{:.0f}'.format(size))
#     else:
#         size = i.LookupParameter('Диаметр').AsDouble() * k
#         sizesOfPipesHolod.append('ø{:.0f}'.format(size))
# #       if systemName.split()[1] in systemDict:
# #           sizesOfPipesHolod.append(systemDict[systemName.split()[1]])
# #       else:
# #           sizesOfPipesHolod.append('Не определено')


def clean_sys_name(sys_name):
    sys_name = sys_name.split('/')[0]
    sys_name = sys_name.replace('Х1.', 'Х.').replace('Х2.', 'Х.').replace('Х3.', 'Х.').replace('Х4.', 'Х.').replace('Хд.', 'Х.')
    return sys_name


for el in pipes:
    sys_name = clean_sys_name(el.LookupParameter('Имя системы').AsString())
    el.LookupParameter('ХТ Имя системы').Set(sys_name)

    size = el.LookupParameter('Диаметр').AsDouble() * k
    if size in pipesDict:
        size = pipesDict[size]
    else:
        size = 'ø{:.0f}'.format(size)
    el.LookupParameter('ХТ Размер фитинга ОВ').Set(size)




 # -------------------------------------- Фитинги и арматура труб -------------------------
 # -------------------------------------- Фитинги и арматура труб -------------------------
 # -------------------------------------- Фитинги и арматура труб -------------------------
 # -------------------------------------- Фитинги и арматура труб -------------------------
 # -------------------------------------- Фитинги и арматура труб -------------------------
 # -------------------------------------- Фитинги и арматура труб -------------------------
 # -------------------------------------- Фитинги и арматура труб -------------------------
 # -------------------------------------- Фитинги и арматура труб -------------------------
 # -------------------------------------- Фитинги и арматура труб -------------------------
 # -------------------------------------- Фитинги и арматура труб -------------------------
 # -------------------------------------- Фитинги и арматура труб -------------------------
 # -------------------------------------- Фитинги и арматура труб -------------------------
 # -------------------------------------- Фитинги и арматура труб -------------------------
 # -------------------------------------- Фитинги и арматура труб -------------------------
 # -------------------------------------- Фитинги и арматура труб -------------------------
 # -------------------------------------- Фитинги и арматура труб -------------------------
 # -------------------------------------- Фитинги и арматура труб -------------------------
 # -------------------------------------- Фитинги и арматура труб -------------------------


# # sizesOfFits = []  # -------------------------------------- Фитинги и арматура труб -------------------------
# # systemNamesOfFits = []
# # LenOfFits = []
# for pFit in pFits:
#     systemName = 'Не определено'
#     originalSystemName = pFit.LookupParameter('Имя системы').AsString()
#     if originalSystemName:
#         systemName = originalSystemName.split('/')[0].replace('обр', '').replace('под', '').replace('Xо ', 'X').replace('Xп ', 'X').replace('Xд ', 'X')
#         systemName = systemName.replace('Х3.', 'Х.').replace('Х4.', 'Х.').replace('Хд.', 'Х.')
#     pFit.LookupParameter('ХТ Имя системы').Set(systemName)


#     # if 'T1 ' in systemName or 'T2 ' in systemName:  # Латинская t
#     #     if systemName.split()[1] in systemOtherDict:
#     #         sizesOfFits.append(systemOtherDict[systemName.split()[1]])
#     #     else:
#     #         sizesOfFits.append('Не определено0')
#     # elif 'Х1 ' in systemName or 'Х2 ' in systemName:
#     #     if systemName.split()[1] in systemDict:
#     #         sizesOfFits.append(systemDict[systemName.split()[1]])
#     #     else:
#     #         sizesOfFits.append('Не определено1')
#     if 'Не учитывать' not in originalSystemName:
#         size = pFit.LookupParameter('Размер').AsString()
#         # if doc.GetElement(pFit.GetTypeId()).LookupParameter('Описание').AsString() and 'ройник' in doc.GetElement(pFit.GetTypeId()).LookupParameter('Описание').AsString():
#         if 'ройник' in pFit.LookupParameter('Тип').AsValueString():
#             size = 'ø' + size.replace(',00 мм', '')
#             # sizesOfFits.append(size)
#             way = 1
#         else:
#             # if size:
#                 size = size.split()[0].replace(',', '.')
#                 if '-' in size:
#                     arr = size.split('-')
#                     if len(arr) == 2:
#                         if arr[0] == arr[1]:
#                             size = pipesDict[float(arr[0])] if float(arr[0]) in pipesDict else 'ø' + size
#                         else:
#                             part0 = pipesDict[float(arr[0])] if float(arr[0]) in pipesDict else arr[0]
#                             part1 = pipesDict[float(arr[1])] if float(arr[1]) in pipesDict else arr[1]
#                             size = '{} - {}'.format(part0, part1)
#                     elif len(arr) == 3:
#                             if arr[0] == arr[1] == arr[2]:
#                                 size = pipesDict[float(arr[0])] if float(arr[0]) in pipesDict else 'ø' + size
#                             else:
#                                 part0 = pipesDict[float(arr[0])] if float(arr[0]) in pipesDict else arr[0]
#                                 part1 = pipesDict[float(arr[1])] if float(arr[1]) in pipesDict else arr[1]
#                                 part2 = pipesDict[float(arr[2])] if float(arr[2]) in pipesDict else arr[2]
#                                 size = '{} - {} - {}'.format(part0, part1, part2)

#                     # try:
#                     #     size = '{}-{}'.format(pipesDict[size.split('-')[0]] + pipesDict[size.split('-')[1]])
#                     #     way = 2
#                     # except KeyError:
#                     #     way = 3
#                     #     if size.split('-')[0] == size.split('-')[1]:
#                     #         way = 4
#                     #         size = size.split('-')[0]
#                 else:
#                     # raise Exception('Дописать тут')

#                     if size in pipesDict:
#                     #     way = 5
#                         size = pipesDict[size]
#                     else:
#                         size = float(size)
#                     #     way = 6
#                     size = 'ø{:.0f}'.format(size)
#     else:
#         size = 'Не определено'
#     pFit.LookupParameter('ХТ Размер фитинга ОВ').Set(size)


#     length = 0
#     param = pFit.LookupParameter('Длина')
#     if param:
#         length = round(param.AsDouble() * 1.1, 2)
#     pFit.LookupParameter('ХТ Длина ОВ').Set(length)


# # for i, pos in enumerate(pFits):
#     # pos.LookupParameter('ХТ Размер фитинга ОВ').Set(sizesOfFits[i])
#     # pos.LookupParameter('ХТ Имя системы').Set(systemNamesOfFits[i])
#     # pos.LookupParameter('ХТ Длина ОВ').Set(LenOfFits[i])



# armsHolod = []
# armsOther = []
# for i in pArms:
#     if i.LookupParameter('Имя системы').AsString():
#         system = i.LookupParameter('Имя системы').AsString().split('/')[0]
#         # if 'Х' in system:           # буква Ха, не Икс (он на меди)
#         if 'Х' in system and 'Х1.' not in system and 'Х2.' not in system and 'Х3.' not in system and 'Х4.' not in system:  # Кириллическая Ха
#             armsHolod.append(i)
#         else:
#             armsOther.append(i)

# armsOtherSystemName = []
# sizesOfArmsOther = []
# for i in armsOther:
#     systemName = i.LookupParameter('Имя системы').AsString().split('/')[0]
#     armsOtherSystemName.append(systemName)
#     name = doc.GetElement(i.GetTypeId()).get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString()
#     if 'Клапан трехходовой с приводом' in name:
#         sizesOfArmsOther.append(doc.GetElement(i.GetTypeId()).get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString())
#     elif 'Насос циркуляционный' in name:
#         sizesOfArmsOther.append(doc.GetElement(i.GetTypeId()).get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString())
#     elif '_Манометры-ТМ-серии-10_ТМ-510' in name:
#         sizesOfArmsOther.append(doc.GetElement(i.GetTypeId()).get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString())
#     elif 'Кран трёхходовой для манометра' in name:
#         sizesOfArmsOther.append(doc.GetElement(i.GetTypeId()).get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString())
#     elif '_Термометр_БТ-32.211' in name:
#         sizesOfArmsOther.append(doc.GetElement(i.GetTypeId()).get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString())
#     elif 'ADSK_АвтоматическийВоздухоудалитель' in name:
#         sizesOfArmsOther.append(doc.GetElement(i.GetTypeId()).get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString())
#     elif 'Датчик температурыrfa' in name:
#         sizesOfArmsOther.append(' ')
#     elif 'Кран шаровой сливной' in name:
#         sizesOfArmsOther.append(doc.GetElement(i.GetTypeId()).get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString())
#     else:
#         if 'T1 ' in systemName or 'T2 ' in systemName:  # Латинская t
#             if systemName.split()[1] in systemOtherDict:
#                 sizesOfArmsOther.append(systemOtherDict[systemName.split()[1]])
#             else:
#                 sizesOfArmsOther.append('Не определено')

#         else:
#             size = i.LookupParameter('Размер').AsString()
#             size = size.split()[0].replace(',', '.')
#             size = float(size)
#             if size in pipesDict:
#                 sizesOfArmsOther.append(pipesDict[size])
#             else:
#                 sizesOfArmsOther.append('ø{:.0f}'.format(size))


# armsHolodSystemName = []
# sizesOfArmsHolod = []
# for i in armsHolod:
#     systemName = i.LookupParameter('Имя системы').AsString().split('/')[0]
#     armsHolodSystemName.append(systemName)
#     name = doc.GetElement(i.GetTypeId()).get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString()
#     if 'Клапан трехходовой с приводом' in name:
#         sizesOfArmsHolod.append(doc.GetElement(i.GetTypeId()).get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString())
#     elif 'Насос циркуляционный' in name:
#         sizesOfArmsHolod.append(doc.GetElement(i.GetTypeId()).get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString())
#     elif '_Манометры-ТМ-серии-10_ТМ-510' in name:
#         sizesOfArmsHolod.append(doc.GetElement(i.GetTypeId()).get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString())
#     elif 'Кран трёхходовой для манометра' in name:
#         sizesOfArmsHolod.append(doc.GetElement(i.GetTypeId()).get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString())
#     elif '_Термометр_БТ-32.211' in name:
#         sizesOfArmsHolod.append(doc.GetElement(i.GetTypeId()).get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString())
#     elif 'ADSK_АвтоматическийВоздухоудалитель' in name:
#         sizesOfArmsHolod.append(doc.GetElement(i.GetTypeId()).get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString())
#     elif 'Датчик температурыrfa' in name:
#         sizesOfArmsHolod.append(' ')
#     elif 'Кран шаровой сливной' in name:
#         sizesOfArmsHolod.append(doc.GetElement(i.GetTypeId()).get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString())
#     else:
#         if 'Х1' == systemName or 'Х2' == systemName:  # Системы магистальных участков
#             print(i.Id)
#             if i.LookupParameter('Диаметр'):
#                 size = i.LookupParameter('Диаметр').AsDouble() * k
#             else:
#                 size = i.LookupParameter('Условный диаметр').AsDouble() * k
#             sizesOfArmsHolod.append('ø{:.0f}'.format(size))
#         else:
#             if systemName.split()[1] in systemDict:
#                 sizesOfArmsHolod.append(systemDict[systemName.split()[1]])
#             else:
#                 sizesOfArmsHolod.append('Не определено')



for el in list(pFits) + list(pArms):
    sys_name = clean_sys_name(el.LookupParameter('Имя системы').AsString())
    el.LookupParameter('ХТ Имя системы').Set(sys_name)

    size = el.LookupParameter('Размер').AsString().replace(',', '.').split('-')
    size_list = sorted([float(i) for i in set(size)], reverse=True)
    done_sizes = []
    for size in size_list:
        if size in pipesDict:
            size = pipesDict[size]
        else:
            size = 'ø{:.0f}'.format(size)
        done_sizes.append(size)
    komms = doc.GetElement(el.GetTypeId()).LookupParameter('Комментарии к типоразмеру').AsString()
    komms = komms or ''

    el.LookupParameter('ХТ Размер фитинга ОВ').Set('{} {}'.format(komms, '-'.join(done_sizes)))

    if el.LookupParameter('ХТ Длина ОВ'):
        param = el.LookupParameter('Длина')
        if param:
            'Проверить единицы измерения параметра Длина, если это "Длина", а не "Число", то добавить тут коэффициент перевода единиц "k"'
        length = round(param.AsDouble() * 1.1, 2) if param else 0
        # print(el.Id)
        el.LookupParameter('ХТ Длина ОВ').Set(length)


# ---------------------------------------------------------------------------------------------------------------
# ---------------------------------------------------------------------------------------------------------------
# ---------------------------------------------------------------------------------------------------------------
# ---------------------------------------------------------------------------------------------------------------
# ---------------------------------------------------------------------------------------------------------------
# ---------------------------------------------------------------------------------------------------------------
# ---------------------------------------------------------------------------------------------------------------
# ---------------------------------------------------------------------------------------------------------------
# ---------------------------------------------------------------------------------------------------------------
# ---------------------------------------------------------------------------------------------------------------
# ---------------------------------------------------------------------------------------------------------------
# ---------------------------------------------------------------------------------------------------------------
# ---------------------------------------------------------------------------------------------------------------
# ---------------------------------------------------------------------------------------------------------------

equipmentSystems = []
for i in equipments:
    name = i.LookupParameter('Имя системы').AsString()
    if name:
        # equipmentSystems.append(name.split('/')[0].split(',')[0].replace('обр', '').replace('под', '').replace('Xо ', 'X').replace('Xп ', 'X').replace('Xд ', 'X').replace('Х3.', 'Х.').replace('Х4.', 'Х.').replace('Хд.', 'Х.'))
        name = clean_sys_name(name)
        equipmentSystems.append(name.split(',')[0])
    else:
        equipmentSystems.append('')


for piping_system in piping_systems:
    name = piping_system.LookupParameter('Имя системы').AsString()
    name = name.split('/')[0].split(',')[0].replace('обр', '').replace('под', '').replace('Xо ', 'X').replace('Xп ', 'X').replace('Xд ', 'X')
    if piping_system.LookupParameter('ХТ Имя системы'):
        piping_system.LookupParameter('ХТ Имя системы').Set(name)


output = script.get_output()

equipmentsSizes = []
for i in equipments:
    typeMarka = doc.GetElement(i.GetTypeId()).LookupParameter('Маркировка типоразмера').AsString()
    if typeMarka:
        link = output.linkify(i.Id, i.LookupParameter('Имя системы').AsString() + ': ' + i.Name + ': ' + (doc.GetElement(i.GetTypeId()).LookupParameter('Описание').AsString()[:25] if doc.GetElement(i.GetTypeId()).LookupParameter('Описание') else '?'))
        print('Ошибка: не следует использовать "Маркировка типоразмера". Следует применять "Комментарии к типоразмеру" {}'.format(link))
    typeMarka = doc.GetElement(i.GetTypeId()).LookupParameter('Комментарии к типоразмеру').AsString()
    if typeMarka:
        equipmentsSizes.append(typeMarka)
    else:
        equipmentsSizes.append('')

for i, pos in enumerate(rects):
    pos.LookupParameter('ХТ Размер фитинга ОВ').Set(rectsSize[i])  # Размер воздуховодов
for i, pos in enumerate(rounds):
    pos.LookupParameter('ХТ Размер фитинга ОВ').Set(roundsDiameter[i])  # Диаметр воздуховодов (и гибких)
for i, pos in enumerate(flexes):
    pos.LookupParameter('ХТ Размер фитинга ОВ').Set(flexDiameter[i])

for i, pos in enumerate(ducts):
    pos.LookupParameter('ХТ Длина ОВ').Set(lengthDucts[i])  # Длина воздуховодов
for i, pos in enumerate(flexes):
    pos.LookupParameter('ХТ Длина ОВ').Set(lengthFlexes[i])
for i, pos in enumerate(pipes):
    pos.LookupParameter('ХТ Длина ОВ').Set(lengthPipes[i])
for i, pos in enumerate(insuls):
    pos.LookupParameter('ХТ Длина ОВ').Set(areaInsuls[i])
    pos.LookupParameter('ХТ Размер фитинга ОВ').Set(insThickness[i])
    th = 10 / k if insThickness[i] == 'δ=10' else 32 / k
    pos.LookupParameter('Толщина изоляции').Set(th)

for i, pos in enumerate(list(dFits) + list(dArms)):
    pos.LookupParameter('ХТ Размер фитинга ОВ').Set(sizes[i])  # Размер фитингов и арматуры воздуховодов

for i, pos in enumerate(list(dFits) + list(terms) + list(dArms) + list(ducts) + list(flexes) + list(insuls)):
    pos.LookupParameter('ХТ Имя системы').Set(systemNamesVent[i])
    comms = pos.LookupParameter('Комментарии').AsString()
    if comms:
        if 'Под вопросом' in comms:
            pos.LookupParameter('ХТ Имя системы').Set(systemNamesVent[i] + ' под вопросом')

############################################################ for i, pos in enumerate(pipesOther):
############################################################     pos.LookupParameter('ХТ Размер фитинга ОВ').Set(sizesOfPipesOther[i])
############################################################     pos.LookupParameter('ХТ Имя системы').Set(systemNamesOther[i])
############################################################ for i, pos in enumerate(pipesHolod):
############################################################     pos.LookupParameter('ХТ Размер фитинга ОВ').Set(sizesOfPipesHolod[i])
############################################################     pos.LookupParameter('ХТ Имя системы').Set(systemNamesHolod[i])
###########################################################
############################################################ # for i, pos in enumerate(pFits):
############################################################ #     pos.LookupParameter('ХТ Размер фитинга ОВ').Set(sizesOfFits[i])
############################################################ #     pos.LookupParameter('ХТ Имя системы').Set(systemNamesOfFits[i])
############################################################ #     pos.LookupParameter('ХТ Длина ОВ').Set(LenOfFits[i])
###########################################################
############################################################ for i, pos in enumerate(armsOther):
############################################################     pos.LookupParameter('ХТ Размер фитинга ОВ').Set(sizesOfArmsOther[i])
############################################################     pos.LookupParameter('ХТ Имя системы').Set(armsOtherSystemName[i])
############################################################ for i, pos in enumerate(armsHolod):
############################################################     pos.LookupParameter('ХТ Размер фитинга ОВ').Set(sizesOfArmsHolod[i])
############################################################     pos.LookupParameter('ХТ Имя системы').Set(armsHolodSystemName[i])

for i, pos in enumerate(equipments):
    if equipmentSystems[i] != '':
        pos.LookupParameter('ХТ Имя системы').Set(equipmentSystems[i])
    if equipmentsSizes[i] != '':
        pos.LookupParameter('ХТ Размер фитинга ОВ').Set(equipmentsSizes[i])
    if pos.LookupParameter('Категория').AsValueString() == 'Сантехнические приборы':
        # print(11)
        # if 'В1' in pos.LookupParameter('Имя системы').AsString():
            # print(222)
        pos.LookupParameter('ХТ Имя системы').Set('В1')


for i in equipments:
    if i.GetSubComponentIds():
        # print('---------------')
        # print(i.Id)
        sys_name = i.LookupParameter('ХТ Имя системы').AsString() or ''
        # print(sys_name)
        sys_name = clean_sys_name(sys_name)
        # print(sys_name)
        for sub_el_id in i.GetSubComponentIds():
            doc.GetElement(sub_el_id).LookupParameter('ХТ Имя системы').Set(sys_name)

for i, pos in enumerate(terms):
    pos.LookupParameter('ХТ Размер фитинга ОВ').Set(sizesOfTerms[i])

# TransactionManager.Instance.TransactionTaskDone()

insulsDictT = {  # -------------------------- Изоляция труб -------------------
    15: '⌀=22, δ=19',
    20: '⌀=28, δ=25',
    25: '⌀=35, δ=25',
    32: '⌀=42, δ=32',
    40: '⌀=48, δ=32',
    50: '⌀=60, δ=40',
    65: '⌀=76, δ=40',
    80: '⌀=89, δ=40'
}
insulsDictX = {
    6.35 : '⌀=6, δ=9',
    9.52 : '⌀=10, δ=9',
    12.7 : '⌀=12, δ=9',
    15.9 : '⌀=15, δ=13',
    19.05: '⌀=18, δ=13',
    22.23: '⌀=22, δ=13',
    25.4 : '⌀=25, δ=13',
    28.58: '⌀=28, δ=13',
    34.93: '⌀=35, δ=13',
    41.28: '⌀=42, δ=13',
    15: '⌀=22, δ=9',
    20: '⌀=28, δ=9',
    25: '⌀=35, δ=9',
    32: '⌀=42, δ=9',
    40: '⌀=48, δ=9',
    50: '⌀=60, δ=9',
    65: '⌀=76, δ=9',
    80: '⌀=89, δ=9',
    90: '⌀=102, δ=9',
    100:'⌀=114, δ=9',
    125:'⌀=140, δ=13',
    150:'⌀=160, δ=13'
}

pipeInsulsSizes = []
pipeInsulsLens = []
pipeInsulsSystems = []
for i in pipeInsuls:
    pipe = doc.GetElement(i.HostElementId)
    size = pipe.LookupParameter('ХТ Размер фитинга ОВ').AsString()
    if size.find('" (') == -1:
        size = float(size.replace('ø', '').replace(',', '.'))
    else:
        # print('--------------')
        # print(size)
        # print(size.find('('))
        # print(size[size.find('(') + 1:])
        try:
            size = float(size[size.find('(') + 1:].replace(',', '.').replace(' mm)', ''))
        except ValueError:
            print('Предупреждение при обработке изоляции перехода(?)')
            size = float(size[size.find('(') + 1:].replace(',', '.').split(' mm)')[0])
    if pipe.LookupParameter('ХТ Имя системы'):
        if pipe.LookupParameter('ХТ Имя системы').AsString():
            system = pipe.LookupParameter('ХТ Имя системы').AsString()
        else:
            system = 'line 505: not string'
    else:
        system = 'line 507: not parameter'
    pipeInsulsSystems.append(system)
    if 'T' in system:  # Для теплоснабжения
        if size in insulsDictT:
            insuls_ = insulsDictT[size]
        else:
            insuls_ = 'Error.T: ' + str(size)
    elif 'X' in system or 'Х' in system:  # Для холодоснабжения
        if size in insulsDictX:
            insuls_ = insulsDictX[size]
        else:
            insuls_ = 'Error.X: ' + str(size)
    else:
        insuls_ = 'Error7: ' + system
    pipeInsulsSizes.append(insuls_)
    pipeInsulsLens.append(i.LookupParameter('Длина').AsDouble() * 1.1)

###################################### Фейки #####     ######     ######     ######     ######     ######
###################################### Фейки #####     ######     ######     ######     ######     ######
###################################### Фейки #####     ######     ######     ######     ######     ######
###################################### Фейки #####     ######     ######     ######     ######     ######
###################################### Фейки #####     ######     ######     ######     ######     ######
###################################### Фейки #####     ######     ######     ######     ######     ######
###################################### Фейки #####     ######     ######     ######     ######     ######
###################################### Фейки #####     ######     ######     ######     ######     ######
###################################### Фейки #####     ######     ######     ######     ######     ######
###################################### Фейки #####     ######     ######     ######     ######     ######
###################################### Фейки #####     ######     ######     ######     ######     ######
###################################### Фейки #####     ######     ######     ######     ######     ######
###################################### Фейки #####     ######     ######     ######     ######     ######
###################################### Фейки #####     ######     ######     ######     ######     ######

fakesForBrackets = []
fakesForArea = []
fakesForPaint = []
fakesForAreaPrimer = []
fakesForPrimer = []
fakesForCheckuot = []
for i in fakes:
    nameOfType = doc.GetElement(i.GetTypeId()).get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    if 'Фейк для Кронш' in nameOfType: fakesForBrackets.append(i)
    elif 'Фейк для площади' in nameOfType: fakesForArea.append(i)
    elif 'Фейк для краски' in nameOfType: fakesForPaint.append(i)
    elif 'Фейк для грунтования площади' in nameOfType: fakesForAreaPrimer.append(i)
    elif 'Фейк для грунтовки' in nameOfType: fakesForPrimer.append(i)
    elif 'Фейк для испытаний давлением' in nameOfType: fakesForCheckuot.append(i)

uniqueSystemAndSize = []
uniqueSystems = []
for i in pipes:
    if i.LookupParameter('Имя системы').AsString():
        system = i.LookupParameter('Имя системы').AsString().split('/')[0].replace('обр', '').replace('под', '').replace('Xо ', 'X').replace('Xп ', 'X').replace('Xд ', 'X')
        size = i.LookupParameter('ХТ Размер фитинга ОВ').AsString()
        type_ = (system, size)
        if type_ not in uniqueSystemAndSize:  # Латинская буква Икс - фреоновые медные трубы
            uniqueSystemAndSize.append(type_)
        if system not in uniqueSystems and 'X' not in system and 'ПУ' not in system:
            uniqueSystems.append(system)

uniqueSystemAndSizeLens = [0 for i in range(len(uniqueSystemAndSize))]
uniqueSystemsAreas = [0 for i in range(len(uniqueSystems))]

outerDiameterDict = {
    6: 6 + 2 * 2.2,
    10: 10 + 2 * 2.2,
    15: 15 + 2 * 2.8,
    16: 16,  # 14.10.2019 для трубки дренажа
    19: 19,  # 14.10.2019
    35: 35,  # 14.10.2019
    20: 20 + 2 * 2.8,
    25: 25 + 2 * 3.2,
    32: 32 + 2 * 3.2,
    40: 40 + 2 * 3.5,
    50: 50 + 2 * 3.5,
    65: 65 + 2 * 4.0,
    80: 80 + 2 * 4.0,
    90: 90 + 2 * 4.0,
    100: 100 + 2 * 4.5,
    125: 125 + 2 * 4.5,
    150: 150 + 2 * 4.5
}

warning = False
for i in pipes:
    if i.LookupParameter('Имя системы').AsString():
        system = i.LookupParameter('Имя системы').AsString().split('/')[0].replace('обр', '').replace('под', '').replace('Xо ', 'X').replace('Xп ', 'X').replace('Xд ', 'X')
    else:
        system = '-'
    size = i.LookupParameter('ХТ Размер фитинга ОВ').AsString()
    type_ = (system, size)
    if type_ in uniqueSystemAndSize:
        l = i.LookupParameter('ХТ Длина ОВ').AsDouble() * k
        uniqueSystemAndSizeLens[uniqueSystemAndSize.index(type_)] += l
    if system in uniqueSystems:
        l = i.LookupParameter('ХТ Длина ОВ').AsDouble() * k
        d = i.LookupParameter('ХТ Размер фитинга ОВ').AsString()
        try:
            d = float(d.replace('ø', ''))
        except ValueError:
            pass
        if d in outerDiameterDict:
            area = 3.14159 * outerDiameterDict[d] * l / 1000 / k
            uniqueSystemsAreas[uniqueSystems.index(system)] += area
        else:
            # raise Exception('outerDiameterDict is not full, require add more position | ' + str(d))
            # print('outerDiameterDict is not full, require add more position | ' + str(d))
            warning = True
if warning:
    print('При расчёте фейков возникла ошибка (Не рассчитаны Кронштейны, краски, грунтовки, испытания)')


fakesForBracketsAmount = [math.ceil(i / 2000) * 1000 for i in uniqueSystemAndSizeLens]
fakesForBracketsSystems = [i for (i, j) in uniqueSystemAndSize]
fakesForBracketsSizes = [j for (i, j) in uniqueSystemAndSize]

paints = [i * 0.11 * 2 for i in uniqueSystemsAreas]

primers = [i * 0.11 for i in uniqueSystemsAreas]

################################## Площадь дактов ########################################

uniqueSystemAndSizeDuctsElements = []
uniqueSystemAndSizeDucts = []
for i in ducts:
    sys_name = i.LookupParameter('Имя системы').AsString()
    sys_name = sys_name if sys_name else 'Не определено'
    system = sys_name.split('/')[0]
    size = i.LookupParameter('ХТ Размер фитинга ОВ').AsString()
    type_ = (system, size)
    if type_ not in uniqueSystemAndSizeDucts:  # Латинская буква Икс - фреоновые медные трубы
        uniqueSystemAndSizeDuctsElements.append(i)
        uniqueSystemAndSizeDucts.append(type_)

uniqueSystemAndSizeDuctsAreas = [0 for i in range(len(uniqueSystemAndSizeDucts))]
for i in ducts:
    sys_name = i.LookupParameter('Имя системы').AsString()
    sys_name = sys_name if sys_name else 'Не определено'
    system = sys_name.split('/')[0]
    size = i.LookupParameter('ХТ Размер фитинга ОВ').AsString()
    type_ = (system, size)
    if type_ in uniqueSystemAndSizeDucts:
        a = i.LookupParameter('Площадь').AsDouble() / kkk * k / 1000 * 1.1
        uniqueSystemAndSizeDuctsAreas[uniqueSystemAndSizeDucts.index(type_)] += a

uniqueSystemAndSizeDuctsAreasForEach = []
for i in ducts:
    sys_name = i.LookupParameter('Имя системы').AsString()
    sys_name = sys_name if sys_name else 'Не определено'
    system = sys_name.split('/')[0]
    size = i.LookupParameter('ХТ Размер фитинга ОВ').AsString()
    type_ = (system, size)
    uniqueSystemAndSizeDuctsAreasForEach.append(uniqueSystemAndSizeDuctsAreas[uniqueSystemAndSizeDucts.index(type_)])

########################################
uniqueSystemAndSizeFlexes = []
for i in flexes:
    sys_name = i.LookupParameter('Имя системы').AsString()
    sys_name = sys_name if sys_name else 'Не определено'
    system = sys_name.split('/')[0]
    size = i.LookupParameter('ХТ Размер фитинга ОВ').AsString()
    type_ = (system, size)
    if type_ not in uniqueSystemAndSizeDucts:  # Латинская буква Икс - фреоновые медные трубы
        uniqueSystemAndSizeFlexes.append(type_)

uniqueSystemAndSizeFlexesAreas = [0 for i in range(len(uniqueSystemAndSizeFlexes))]
for i in flexes:
    sys_name = i.LookupParameter('Имя системы').AsString()
    sys_name = sys_name if sys_name else 'Не определено'
    system = sys_name.split('/')[0]
    size = i.LookupParameter('ХТ Размер фитинга ОВ').AsString()
    type_ = (system, size)
    if type_ in uniqueSystemAndSizeFlexes:
        a = i.LookupParameter('Длина').AsDouble() * 1.1 * 3.14 * i.LookupParameter('Диаметр').AsDouble() / kk
        uniqueSystemAndSizeFlexesAreas[uniqueSystemAndSizeFlexes.index(type_)] += a

uniqueSystemAndSizeFlexesAreasForEach = []
for i in flexes:
    sys_name = i.LookupParameter('Имя системы').AsString()
    sys_name = sys_name if sys_name else 'Не определено'
    system = sys_name.split('/')[0]
    size = i.LookupParameter('ХТ Размер фитинга ОВ').AsString()
    type_ = (system, size)
    uniqueSystemAndSizeFlexesAreasForEach.append(uniqueSystemAndSizeFlexesAreas[uniqueSystemAndSizeFlexes.index(type_)])

########################################
uniqueSystemAndSizeDFits = []
for i in dFits:
    sys_name = i.LookupParameter('Имя системы').AsString()
    sys_name = sys_name if sys_name else 'Не определено'
    system = sys_name.split('/')[0]
    size = i.LookupParameter('ХТ Размер фитинга ОВ').AsString()
    naim = doc.GetElement(i.GetTypeId()).LookupParameter('Описание').AsString()
    type_ = (system, size, naim)
    if type_ not in uniqueSystemAndSizeDFits:  # Латинская буква Икс - фреоновые медные трубы
        uniqueSystemAndSizeDFits.append(type_)

uniqueSystemAndSizeDFitsAreas = [0 for i in range(len(uniqueSystemAndSizeDFits))]
for i in dFits:
    sys_name = i.LookupParameter('Имя системы').AsString()
    sys_name = sys_name if sys_name else 'Не определено'
    system = sys_name.split('/')[0]
    size = i.LookupParameter('ХТ Размер фитинга ОВ').AsString()
    naim = doc.GetElement(i.GetTypeId()).LookupParameter('Описание').AsString()
    type_ = (system, size, naim)
    if type_ in uniqueSystemAndSizeDFits:
        if i.LookupParameter('ХТ Площадь'):
            a = i.LookupParameter('ХТ Площадь').AsDouble() / kk
        else:
            a = 0.0
        uniqueSystemAndSizeDFitsAreas[uniqueSystemAndSizeDFits.index(type_)] += a

uniqueSystemAndSizeDFitsAreasForEach = []
for i in dFits:
    sys_name = i.LookupParameter('Имя системы').AsString()
    sys_name = sys_name if sys_name else 'Не определено'
    system = sys_name.split('/')[0]
    size = i.LookupParameter('ХТ Размер фитинга ОВ').AsString()
    naim = doc.GetElement(i.GetTypeId()).LookupParameter('Описание').AsString()
    type_ = (system, size, naim)
    uniqueSystemAndSizeDFitsAreasForEach.append(uniqueSystemAndSizeDFitsAreas[uniqueSystemAndSizeDFits.index(type_)])


if len(fakesForBrackets) != len(uniqueSystemAndSize) or \
len(fakesForArea) != len(uniqueSystems) or \
len(fakesForPaint) != len(uniqueSystems) or \
len(fakesForAreaPrimer) != len(uniqueSystems) or \
len(fakesForPrimer) != len(uniqueSystems) or \
len(fakesForCheckuot) != len(uniqueSystemAndSize):
    # raise Exception('{}, {}, {}, {}, {}, {}'.format(
    print('{}, {}, {}, {}, {}, {}'.format(
        -len(fakesForBrackets) + len(uniqueSystemAndSize),
        -len(fakesForArea) + len(uniqueSystems),
        -len(fakesForPaint) + len(uniqueSystems),
        -len(fakesForAreaPrimer) + len(uniqueSystems),
        -len(fakesForPrimer) + len(uniqueSystems),
        -len(fakesForCheckuot) + len(uniqueSystemAndSize)
        )
    )
else:
    # TransactionManager.Instance.EnsureInTransaction(doc)

    for i, pos in enumerate(fakesForBrackets):
        pos.LookupParameter('ХТ Размер фитинга ОВ').Set(fakesForBracketsSizes[i])
        pos.LookupParameter('ХТ Длина ОВ').Set(fakesForBracketsAmount[i] / k)
        pos.LookupParameter('ХТ Имя системы').Set(fakesForBracketsSystems[i])

    for i, pos in enumerate(fakesForArea):
        pos.LookupParameter('ХТ Размер фитинга ОВ').Set('В два слоя')
        pos.LookupParameter('ХТ Длина ОВ').Set(uniqueSystemsAreas[i])
        pos.LookupParameter('ХТ Имя системы').Set(uniqueSystems[i])

    for i, pos in enumerate(fakesForPaint):
        pos.LookupParameter('ХТ Размер фитинга ОВ').Set('БТ-577')
        pos.LookupParameter('ХТ Длина ОВ').Set(paints[i])
        pos.LookupParameter('ХТ Имя системы').Set(uniqueSystems[i])
        pos.LookupParameter('ADSK_Примечание').Set('В два слоя, S={:.1f}'.format(uniqueSystemsAreas[i] / 3.25).replace('.', ',') + ' м²')

    for i, pos in enumerate(fakesForAreaPrimer):
        pos.LookupParameter('ХТ Размер фитинга ОВ').Set('В один слой')
        pos.LookupParameter('ХТ Длина ОВ').Set(uniqueSystemsAreas[i])
        pos.LookupParameter('ХТ Имя системы').Set(uniqueSystems[i])

    for i, pos in enumerate(fakesForPrimer):
        pos.LookupParameter('ХТ Размер фитинга ОВ').Set('ГФ-021')
        pos.LookupParameter('ХТ Длина ОВ').Set(primers[i])
        pos.LookupParameter('ХТ Имя системы').Set(uniqueSystems[i])
        pos.LookupParameter('ADSK_Примечание').Set('S={:.1f}'.format(uniqueSystemsAreas[i] / 3.25).replace('.', ',') + ' м²')

    for i, pos in enumerate(fakesForCheckuot):
        pos.LookupParameter('ХТ Размер фитинга ОВ').Set(fakesForBracketsSizes[i])
        pos.LookupParameter('ХТ Длина ОВ').Set(uniqueSystemAndSizeLens[i] / k)
        pos.LookupParameter('ХТ Имя системы').Set(fakesForBracketsSystems[i])

for i, pos in enumerate(pipeInsuls):
    pos.LookupParameter('ХТ Размер фитинга ОВ').Set(pipeInsulsSizes[i])
    if 'δ=' in pipeInsulsSizes[i]:
        th = float(pipeInsulsSizes[i].split('δ=')[-1]) / k
        pos.LookupParameter('Толщина изоляции').Set(th)
    # print(pipeInsulsSizes[i])
    # pos.LookupParameter('Толщина изоляции').Set(int(pipeInsulsSizes[i].split('δ=')[1] if pipeInsulsSizes[i].split('δ=') else 1) / k)
    pos.LookupParameter('ХТ Длина ОВ').Set(pipeInsulsLens[i])
    pos.LookupParameter('ХТ Имя системы').Set(pipeInsulsSystems[i])
#       pos.LookupParameter('ХТ Корпус').Set(pipeInsulsKorpus[i])

for i, pos in enumerate(ducts):
    pos.LookupParameter('ADSK_Примечание').Set('S={:.1f}'.format(uniqueSystemAndSizeDuctsAreasForEach[i]).replace('.', ',') + ' м²')
for i, pos in enumerate(flexes):
    pos.LookupParameter('ADSK_Примечание').Set('S={:.1f}'.format(uniqueSystemAndSizeFlexesAreasForEach[i]).replace('.', ',') + ' м²')
for i, pos in enumerate(dFits):
    if uniqueSystemAndSizeDFitsAreasForEach[i] == 0:
        pos.LookupParameter('ADSK_Примечание').Set('')
    else:
        pos.LookupParameter('ADSK_Примечание').Set(('S={:.1f}'.format(uniqueSystemAndSizeDFitsAreasForEach[i]) if uniqueSystemAndSizeDFitsAreasForEach[i] >= 0.1 else 'S={:.2f}'.format(uniqueSystemAndSizeDFitsAreasForEach[i])).replace('.', ',') + ' м²')

# print('---------------ducts')
# for i in list(ducts):
#     print(i)
# print('---------------dFits')
# for i in list(dFits):
#     print(i)
# print('---------------terms')
# for i in list(terms):
#     print(i)
# print('---------------flexes')
# for i in list(flexes):
#     print(i)
# print('---------------dArms')
# for i in list(dArms):
#     print(i)
# print('---------------insuls')
# for i in list(insuls):
#     print(i)
# print('---------------pipes')
# for i in list(pipes):
#     print(i)
# print('---------------pFits')
# for i in list(pFits):
#     print(i)
# print('---------------pArms')
# for i in list(pArms):
#     print(i)
# print('---------------pipeInsuls')
# for i in list(pipeInsuls):
#     print(i)
# print('---------------equipments')
# for i in list(equipments):
#     print(i)
# print('---------------fakes')
# for i in list(fakes):
#     print(i)

all = list(ducts) + list(dFits) + list(terms) + list(flexes) + list(dArms) + list(insuls) + list(pipes) + list(pFits) + list(pArms) + list(pipeInsuls) + list(equipments) + list(fakes)
for i in all:
    if not isinstance(i, str):
        op = rf = ''
        op = doc.GetElement(i.GetTypeId()).LookupParameter('Описание').AsString()
        op = op if op else ''
        if i.LookupParameter('ХТ Размер фитинга ОВ'):
            rf = i.LookupParameter('ХТ Размер фитинга ОВ').AsString()
            rf = rf if rf else ''
        i.LookupParameter('оп+рф').Set(op + ' ' + rf)

for i in all:
    if not isinstance(i, str):
        if i.LookupParameter('ХТ Имя системы'):
            sysName = i.LookupParameter('ХТ Имя системы').AsString()
            if not sysName:
                continue
            if sysName[0] == 'П' or sysName[0] == 'К':
                if sysName[1].isdigit():
                    sort = sysName[1:] + ' 1 Приток'
                elif sysName[:2] == 'ПП':
                    sort = '600 ПП' + sysName[2:]
                elif sysName[:2] == 'ПЕ':
                    sort = '700 ПЕ' + sysName[2:]
                elif sysName[:2] == 'ПД':
                    sort = '800 ПД' + sysName[2:]
                elif sysName[:3] == 'Пер':
                    sort = '900 ПД' + sysName[2:]
            elif sysName[0] == 'В':
                if sysName[1].isdigit():
                    sort = sysName[1:] + ' 2 Вытяжка'
            elif sysName[:2] == 'ДУ':
                sort = '500 ДУ' + sysName[2:]
            elif sysName[:2] == 'ДП':
                sort = '400 Д' + sysName[2:] + ' 1 Приток'
            elif sysName[:2] == 'ДВ':
                sort = '400 Д' + sysName[2:] + ' 2 Вытяжка'
            elif sysName[:2] == 'T1':
                sort = str(1000 + int(sysName.split()[-1] if sysName.split()[-1].isdigit() else '00')) + ' T1'
            elif sysName[:2] == 'T2':
                sort = str(1000 + int(sysName.split()[-1] if sysName.split()[-1].isdigit() else '00')) + ' T2'
                # sort = str(1000 + int(sysName.split()[-1])) + ' T2' + (sysName[2:] if sysName[2:] else '00')
                # sort = '1200 T2' + sysName[2:] if sysName[2:] else '00'
                # print(sysName)
                # print(sort)
            else:
                sort = '9999'

            # sort = sysName[1:] if sysName else ''
            # if sort:
            #     sort += ' 1' if sysName[0] == 'П' else ''
            #     sort += ' 2' if sysName[0] == 'В' else ''
            #     sort += ' 3' if sysName[:2] == 'T1' else ''
            #     sort += ' 4' if sysName[:2] == 'T2' else ''
            # if i.LookupParameter('ХТ Имя системы').AsStr ing() == 'T1':
                # print(sort.replace(' ', '*'))
            # print(i.Id)
            i.LookupParameter('Сортировка строка').Set(sort)
            # i.LookupParameter('Сортировка').Set(sort)
    # else:
        # print('isinstance({}, str)'.format(i))
errors = {}
for i in all:
    if not doc.GetElement(i.GetTypeId()).LookupParameter('Стоимость').HasValue:
        if i.GetTypeId() not in errors:
            errors[i.GetTypeId()] = output.linkify(i.Id)
if errors:
    print('Следует прописать стоимость для сортировки и подсчёта количества:\n{}'.format(' '.join([errors[k] for k in errors])))

for i in all:
    if i.LookupParameter('ХТ Имя системы отступ'):
        if i.LookupParameter('ХТ Имя системы').AsString():
            i.LookupParameter('ХТ Имя системы отступ').Set('‎' + ' ' * 90 + i.LookupParameter('ХТ Имя системы').AsString())

OUT = ['{} {} Фейк для кронштейнов'.format(len(fakesForBrackets), len(uniqueSystemAndSize)),
'{} {} Фейк для площади'.format(len(fakesForArea), len(uniqueSystems)),
'{} {} Фейк для краски'.format(len(fakesForPaint), len(uniqueSystems)),
'{} {} Фейк для грунтования площади'.format(len(fakesForAreaPrimer), len(uniqueSystems)),
'{} {} Фейк для грунтовки'.format(len(fakesForPrimer), len(uniqueSystems)),
'{} {} Фейк для испытаний давлением'.format(len(fakesForCheckuot), len(uniqueSystemAndSize))]

length = sum([len(fakesForBrackets), len(uniqueSystemAndSize),
len(fakesForArea), len(uniqueSystems),
len(fakesForPaint), len(uniqueSystems),
len(fakesForAreaPrimer), len(uniqueSystems),
len(fakesForPrimer), len(uniqueSystems),
len(fakesForCheckuot), len(uniqueSystemAndSize)])
if length:
    for i in OUT:
        print(i)

t.SetName('Расчёт спеки {}, Δ={} с'.format(time.ctime().split()[3], time.time() - startTime))
t.Commit()
