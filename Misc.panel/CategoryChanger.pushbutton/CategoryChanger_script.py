# -*- coding: utf-8 -*-
"""Сменить категорию семейства"""
__title__ = 'Сменить\nкатегорию'
__author__ = 'SG'
import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId
import sys
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

els = [doc.GetElement(id) for id in uidoc.Selection.GetElementIds()]


t = Transaction(doc, '123')
t.Start()

i = 1
for el in els:
	el.SheetNumber = 'q ' + str(i)
	i += 1
# print(els[0].SheetNumber)



# for i in dir(doc.GetElement(els[0].GetTypeId())):
#   print(i)
# for el in els:
#     tp = doc.GetElement(el.GetTypeId())
    # s = tp.LookupParameter('Комментарии к типоразмеру').AsString()
    # tp.LookupParameter('Маркировка типоразмера').Set(s if s else '')
    # tp.LookupParameter('Шифр').Set(0) if tp.LookupParameter('Шифр') else None
    # tp.LookupParameter('JSN').Set('XT')

# print(els[0].Definition.GetFilters())
# for f in list(els[0].Definition.GetFilters()):
    # try:
    # print(f.FieldId)
    # print(els[0].Definition.GetFieldId(int(f.FieldId.ToString())))
    # except:
    #     print(2223)
    #     pass

# dir(list(els[0].Definition.GetFilters())[0])

# for i in dir(list(els[0].Definition.GetFilters())[0]):
#     print(i)

# ids = [2088030, 315622, 1659594, 1659601, 1659602, 2092488, 2128953, 1382990, 1383979, 1384864, 1385758, 1387438, 1659595, 1659596, 2090023, 1659597, 1659598, 2233480, 1659599, 1356464, 34623, 1659603, 1659604, 2093761, 1659605, 1659606, 1659574, 1659607, 1659610, 1659611, 1659612, 2100394, 1659589, 1659616, 1659617, 2107174, 1659613, 2101953, 2104487, 2105757, 1659614, 1659579, 1659580, 1659615, 2108250, 914595, 915514, 939974, 940983, 1659622, 1659623, 1659624, 1659625, 1659581, 1659548, 1659618, 1659619, 1659620, 1659621, 1659531, 1659575, 1659576, 1659570, 1114618, 1659571, 1292399, 680519, 1659626, 2116125, 2117266, 1172923, 1007893, 1418072, 916399, 917244, 912759, 913676, 936117, 937210, 938065, 1659532, 1659552, 1659533, 1659628, 1659534, 1659535, 1659536, 1659627, 1659537, 1659538, 1659578, 1659591, 1659564, 1659540, 1659541, 1659543, 1659544, 1659545, 1350262, 1659546, 1659547, 1659629, 1659630, 1659631, 1659632, 2234887, 1659633, 1346213, 1352427, 1659608, 1659609, 1659569, 1659640, 1659567, 1659582, 1659634, 2118530, 449836, 1659573, 1659568, 1659635, 1659636, 1659566, 1659638, 1659637, 1659555, 2122133, 2123286, 1659557, 1659559, 1659560, 1659542, 1659563, 1659562, 1659641, 1659565, 1659539, 1659561, 2119740, 2120974, 1659577, 1659586, 1659528, 1659585, 1659584, 1659554, 1659553, 1659583, 1310820, 1312294, 1659639, 2124668, 1396104, 1659592, 881080, 2126868, 1303917, 1659551, 1659549, 1171992, 1659550, 1659572]
# com = ['AL1', 'STR.01', 'BB1', 'BB10', 'BB11', 'BB12', 'BB12.1', 'BB13', 'BB14', 'BB15', 'BB16', 'BB17', 'BB2', 'BB4', 'BB4.1', 'BB5', 'BB6', 'BB7', 'BB8', 'BB9', 'L1', 'C1', 'C2', 'CC1', 'CC1', 'CC2', 'СС3', 'CC4', 'CC6', 'CC7', 'CC8', 'CM1', 'CM10', 'CM11', 'CM12', 'CM12', 'CM2', 'CM2', 'CM4', 'CM5', 'CM6', 'CM7', 'CM8', 'CM9', 'Е10', 'L1.8', 'L1.9', 'L1.1', 'L1.2', 'E11', 'E12', 'E13', 'E14', 'E15', 'E2', 'E3', 'E4', 'E5', 'E6', 'E7', 'Е8', 'Е9', 'F1', 'F1.1', 'F2', '0', '0', 'HH1', 'KF1.1', 'KF1.2', 'KP', 'L1', 'L1.02', 'L1.10', 'L1.11', 'L1.3', 'L1.4', 'L1.5', 'L1.6', 'L1.7', 'LL8', 'SA1', 'M1', 'M10', 'M11', 'M12', 'M2', 'M3', 'M4', 'M5', 'Е9', 'M7', 'M8', 'M9', 'MF1', 'MF1.1', 'MF1.2', 'MF2', 'MF2.1', 'MF3.1', 'MF3.2', 'MF4', 'MF5', 'MS1', 'MS2', 'MS3', 'MS5', 'N2', 'N2.1', 'NEW1', 'NEW2', 'O1', 'P1', 'P2', 'E15', 'P4', 'PG1', 'PM', 'RH1', 'RH2', 'RH5', 'RH6', 'RH7', 'RH8', 'RH9', 'S1', 'S10.1', 'S10.2', 'S11', 'S12', 'S13', 'S13.1', 'S14', 'S15', 'S16', 'S17', 'S18', 'S19', 'S2.1', 'S2.2', 'S3', 'S4', 'S5', 'S6', 'S7', 'S8', 'S9', 'SA3', 'SA4', 'SA5', 'SF1', 'SF2', 'PRT', 'T1', 'UV', 'VS1', 'WSH', 'ом03', '', 'СУ', 'СУ', 'ШЛВЖ']

# ids = [1171992, 1356464, 917244, 913676, 1387438, 1384864, 1383979, 1310820, 1172923, 449836, 1396104, 1303917, 881080, 939974, 912759, 916399, 938065, 937210, 936117, 34623, 1007893, 1418072, 1114618, 1382990, 940983, 915514, 914595, 1346213, 1352427, 1659549, 1312294, 1385758, 315622]
# com = ['Стол', 'Шкаф для хранения c выдвижными полками,для хранения парафиновых блоков,парафиновых блоков,с Делителеми - для патоморофологических образцов с Программируемыми жетонами и Считыватель карт RFID', 'Центрифуга медицинская  CM-6MT', 'Термостат твердотельный с таймером TT-2 Термит', 'Термопринтер Zebra GX430t', 'Термоконтейнер транспортировочный MVE SC4/2V', 'Термоконтейнер транспортировочный CryoPod', 'Сканер 2D штрих-кодов Perception HD SBS', 'Рабочее место с компьютером', 'Рабочее место с компьютером', 'Принтер', 'Посудомоечная машина', 'Облучатель-рециркулятор СМ45', 'Набор одноканальных дозаторов переменного объема (0,1 – 2,5 мкл,0,5 – 10 мкл,10 – 100 мкл,100 – 1 000 мкл', 'Мини‐центрифуга‐вортекс  FV-2400', 'Миксер-вортекс LP', 'Магнитный штатив на 96 мест для работы с глубоколуночными планшетами и планшетами для ПЦР', 'Магнитный штатив на 16 мест для работы с пробирками объемом 0,5; 1,5 и 2,0 мл', 'Магнитный штатив на 12 пробирок  1,5 мл', 'Ламинарный бокс', 'Ламинарный бокс', 'Ламинарный бокс', 'Лабораторный морозильник', 'Криогенный сосуд дьюара LAB 10', 'Дозатор электронный 1-канальный варьируемого объема,100-5000 мкл', 'Дозатор 8-канальный,переменного объема,10 – 100 мкл', 'Дозатор 8-канальный,переменного объема, 0,5 – 10 мкл', 'Генератор азота модель Whisper-120 L', 'Генератор азота модель Infinity 1034', 'Бокс биологической безопасности 2 класса Safe 2020', 'Автоматизированное устройство открывания и закрывания пробирок LVL SAFE CAP 96', 'MVE XC 34/18 Система хранения в парах жидкого азота', 'Cтерилизатор настольный паровой']

# for id, co in zip(ids, com):
#     # doc.GetElement(ElementId(id)).LookupParameter('Комментарии к типоразмеру').Set(co)
#     doc.GetElement(ElementId(id)).LookupParameter('Описание').Set(co)

# els = FilteredElementCollector(doc)\
#     .OfCategory(BuiltInCategory.OST_SpecialityEquipment)\
#     .WhereElementIsNotElementType().ToElements()

# els = [doc.GetElement(el.GetTypeId()) for el in els]
# els2 = []
# lst = []
# for el in els:
#     if el.Id.IntegerValue not in lst:
#         lst.append(el.Id.IntegerValue)
#         els2.append(el)

# els2 = list(set(els2))

# for el in els2:
    # for i in dir(el):
    #     print(i)
    # break
    # s00 = el.Id.IntegerValue
    # s0 = el.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    # s1 = el.LookupParameter('JSN').AsString() if el.LookupParameter('JSN') else '-'
    # s2 = el.LookupParameter('Комментарии к типоразмеру').AsString() if el.LookupParameter('Комментарии к типоразмеру') else '-'
    # s3 = el.LookupParameter('Маркировка типоразмера').AsString()
    # s4 = el.LookupParameter('Описание').AsString()
    # print('{}1q1q1q1q1q{}1q1q1q1q1q{}1q1q1q1q1q{}1q1q1q1q1q{}1q1q1q1q1q{}'.format(s00, s0, s1, s2, s3, s4))

t.Commit()
