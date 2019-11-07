# -*- coding: utf-8 -*-
"""Shift + Click откроет файл"""
__title__ = 'Заполнение\nпространств'
__author__ = 'SG'

import os
import sys
import re
import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *

from Autodesk.Revit.DB import SectionType, FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


def natural_sorted(list, key=lambda s: s):
    """
    Sort the list into natural alphanumeric order.
    """
    def get_alphanum_key_func(key):
        convert = lambda text: int(text) if text.isdigit() else text
        return lambda s: [convert(c) for c in re.split('([0-9]+)', key(s))]
    sort_key = get_alphanum_key_func(key)
    return sorted(list, key=sort_key)


path = doc.PathName.split(' ОВ v')[0] + ' ТВБ.txt'
try:
    src = open(path, 'r').read().decode("utf-8")[:-1]
except IOError:
    info = path + '\nЭтот файл должен содержать таблицу (разделители - табуляция).\nПервая строка должна содержать заголовки.\nПоследняя строка должна быть пустой.\n'
    with open(path, 'w') as file:
        file.write(info.encode("utf-8"))
    os.startfile(path)
else:
    if __shiftclick__:
        os.startfile(path)
        sys.exit()


titles = src.split('\n')[0].split('\t')
_number = titles.index('№')
_name = titles.index('Наименование помещения')
_supplyAirFlow = titles.index('Расход приточного воздуха, м³/ч')
_supplySystem = titles.index('Приточная система')
_exhaustSystem = titles.index('Вытяжная система')
_exhaustAirFlow = titles.index('Расход удаляемого воздуха, м³/ч')
_exhaustDevices = titles.index(' Вытяжные  воздухораспределительные устройства')
_supplyDevices = titles.index(' Приточные  воздухораспределительные устройства')
# _sortFactor = titles.index('Сортировка')
# _level = titles.index('Уровень')
_cleanlinessClass = titles.index('Класс чистоты')
_temperature = titles.index('Расчетная температура воздуха в помещении')
_area = titles.index('Площадь, м²')
_height = titles.index('Высота, м')
_volume = titles.index('Объем, м³')
_multiplicityInflow = titles.index('Кратность притока, N/ч')
_multiplicityOutflow = titles.index('Кратность вытяжки, N/ч')
_note = titles.index('Примечание')


raws = [x.split('\t') for x in src.split('\n')][1:]                 # raws
number = [x[_number] for x in raws]                                 # nums
name = [x[_name] for x in raws]
supplyAirFlow = [x[_supplyAirFlow] for x in raws]                   # supplyAir
exhaustAirFlow = [x[_exhaustAirFlow] for x in raws]                 # exhaustAir
supplySystem = [x[_supplySystem] for x in raws]                     # sysP
exhaustSystem = [x[_exhaustSystem] for x in raws]                   # sysV
cleanlinessClass = [x[_cleanlinessClass] for x in raws]             # categ
supplyDevices = [x[_supplyDevices] for x in raws]                   # pritok
exhaustDevices = [x[_exhaustDevices] for x in raws]                 # vityajka
temperature = [x[_temperature] for x in raws]
area = [x[_area] for x in raws]
height = [x[_height] for x in raws]
volume = [x[_volume] for x in raws]
multiplicityInflow = [x[_multiplicityInflow] for x in raws]
multiplicityOutflow = [x[_multiplicityOutflow] for x in raws]
note = [x[_note] for x in raws]


spaces = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElements()

t = Transaction(doc, 'Заполнение пространств')
t.Start()

report = []
existNums = []
for space in sorted(spaces, key=lambda x: x.LookupParameter('Номер').AsString()):
    num = space.LookupParameter('Номер').AsString()
    existNums.append(num)
    if num in number:
        index = number.index(space.LookupParameter('Номер').AsString())
        sA = supplyAirFlow[index]
        eA = exhaustAirFlow[index]
        s = s1 = s2 = ''
        if sA and sA != '0':
            s1 = '{}: +{}'.format(supplySystem[index], sA)
        if eA and eA != '0':
            s2 = '{}: -{}'.format(exhaustSystem[index], eA)
        if s1 and s2:
            s = s1 + ', ' + s2
        else:
            s = (s1 or s2) or s
        s += '\n{}, {}, {}'.format(cleanlinessClass[index], supplyDevices[index], exhaustDevices[index]).replace(' ,', '').replace('\n, ', '\n')  # Доп инфа
        if s[-1] == ' ':
            s = s[:-1]
        if s[-1] == ',':
            s = s[:-1]
        old_s = space.LookupParameter('Комментарии').AsString()
        if old_s != s:
            space.LookupParameter('Комментарии').Set(s)
            report.append([num, old_s.replace('\n', '<br>'), s.replace('\n', '<br>')])
            # report.append([num, old_s.replace('\n', ' ☻ '), s.replace('\n', ' ☻ ')])
    else:
        report.append([num, 'Нет в таблице', ''])
        print('В таблице нет {}'.format(num))

for num in sorted(list(set(number))):
    if num:
        if num not in existNums:
            report.append([num, 'Нет в модели', ''])
            print('В модели нет {}'.format(num))

# data = [
# ['row1', 'data', 'data', 80 ],
# ['row2', 'data', 'data', 45 ],
# ]

natural_sorted(report, key=lambda x: x[0])
report.reverse()
# report.sort(key=lambda x: x[0])

from pyrevit import script
output = script.get_output()

# print(len(report))
# for i in report:
#     print(i)
# print(123123123)

if report:
    output.print_table(
        table_data=report,
        # title="Example Table",
        columns=["№", "Старые данные", "Новые данные", ],
        formats=['', '', '', ],
        # last_line_style='color:red;'
    )










schedules = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Schedules).WhereElementIsNotElementType().ToElements()
schedule = list(filter(lambda x: x.Name == 'Таблица воздушного баланса', schedules))[0]
data = schedule.GetTableData().GetSectionData(SectionType.Body)

for i in range(data.LastRowNumber):
    data.RemoveRow(i + 1)

for num in number:
    data.InsertRow(0)

els = FilteredElementCollector(doc, schedule.Id).ToElements()

for index, el in enumerate(els):
    el.LookupParameter('№').Set(number[index] if number[index] else '‎')
    el.LookupParameter('Наименование помещения').Set(name[index])
    if name[index] == 'Итого ':
        el.LookupParameter('№').Set(number[index] if number[index] else '‎‎‎‎‎‎')
        el.LookupParameter('Наименование помещения').Set('‎' + ' ' * 115 + 'Итого')
    el.LookupParameter('Класс чистоты').Set(cleanlinessClass[index])
    el.LookupParameter('Расчетная температура воздуха в помещении').Set(temperature[index])
    el.LookupParameter('Площадь, м²').Set(area[index])
    el.LookupParameter('Высота, м').Set(height[index])
    el.LookupParameter('Объем, м³').Set(volume[index])
    el.LookupParameter('Кратность притока, N/ч').Set(multiplicityInflow[index])
    el.LookupParameter('Кратность вытяжки, N/ч').Set(multiplicityOutflow[index])
    el.LookupParameter('Расход приточного воздуха, м³/ч').Set(supplyAirFlow[index])
    el.LookupParameter('Расход удаляемого воздуха, м³/ч').Set(exhaustAirFlow[index])
    el.LookupParameter('Приточные  воздухораспределительные устройства').Set(supplyDevices[index])
    el.LookupParameter('Вытяжные  воздухораспределительные устройства').Set(exhaustDevices[index])
    el.LookupParameter('Приточная система').Set(supplySystem[index])
    el.LookupParameter('Вытяжная система').Set(exhaustSystem[index])
    el.LookupParameter('Примечание.').Set(note[index])

t.Commit()











# t = Transaction(doc, 'Заполнение пространств')
# t.Start()

# # в конце строки не должно быть пустоты
# src = """1	Рентгеноперационная	Б		21	50,4	2,6	131			3000	2400	FFD 12/14	Вытяжные шкафы	П1	В1
# 2	Предоперационная	В		21	12,6	2,6	33			300	200	H11	4АПН+ЗКСД	П1	В1
# 3	Санпропускник	Г		21	16,3	2,6	42			250		ДПУ		П1	.
# 3.1	Туалет	Г		21	1,9	2,6	5				50		ДПУ		В2
# 3.2	Душ	Г		21	1,9	2,6	5				75		ДПУ		В2
# 3.3	Коридор	Г		21	4,7	2,6	12				100		ДПУ		В1
# 4	Комната управления	В		21	10,8	2,6	28			250	300	H11	4АПН+ЗКСД	П1	В1
# 5	Наркозная	В		21	11,1	2,6	29			800	800	H11	Вытяжной шкаф	П1	В1
# 6	Шлюз-перекладчик	Г		21	13,2	2,6	34			300	200	H11	4АПН+ЗКСД	П1	В1
# 7	Помещение хозинвентаря и дезсредств	Г		21	4,2	2,6	11				50		ДПУ		В1
# 8	Инструментально-материальная	Г		21	7,5	2,6	20	6	4	120	80	ДПУ	ДПУ	П1	В1
# 9	Моечная	Г		21	9,7	2,6	25	3	4	80	100	ДПУ	ДПУ	П1	В1
# 10	Шлюз чистый	Г		21	5	2,6	13			150		ДПУ		П1	.
# 11	Шлюз грязный	Г		21	4,9	2,6	13				250		ДПУ		В1
# к1	Общебольничный коридор			21							400		ДПУ		В1
# к2	Общебольничный коридор			21							200		ДПУ		В1"""

# spaces = FilteredElementCollector(doc)\
#     .OfCategory(BuiltInCategory.OST_MEPSpaces)\
#     .WhereElementIsNotElementType().ToElements()

# # values = [x.split('\t') for x in src.split('\n')] Практика мэппинга:
# values = map(lambda x: x.split('\t'), src.split('\n'))
# # for i in values:
# #     print('-------------')
# #     for j in i:
# #       if j:
# #           print(j)
# #       else:
# #           print('-')
# nums = map(lambda x: x[0], values)
# # print nums
# names = map(lambda x: x[1], values)
# categ = map(lambda x: x[2], values)
# supplyAir = map(lambda x: x[10], values)
# exhaustAir = map(lambda x: x[11], values)
# pritok = map(lambda x: x[12], values)
# # for i, pos in enumerate(values):
# # 	print '------------------------------------'
# # 	print i
# # 	print pos
# vityajka = map(lambda x: x[13], values)
# sysP = map(lambda x: x[14], values)
# sysV = map(lambda x: x[15], values)
# # vityajka = [i[13] for i in values]
# # vityajka = []
# # for i in values:
# # 	print('{} - {}'.format(i[12], i[13]))
# # 	vityajka.append(i[13])
# # print 1222
# existNums = []
# for sp in sorted(spaces, key=lambda x: x.LookupParameter('Номер').AsString()):
#     n = sp.LookupParameter('Номер').AsString()
#     existNums.append(n)
#     if n in nums:
#         print(n)
#         index = nums.index(sp.LookupParameter('Номер').AsString())
#         sA = supplyAir[index]
#         eA = exhaustAir[index]
#         s = ''
#         s1 = ''
#         s2 = ''
#         # if sA:
#         #     s1 = '+{}'.format(sA)
#         # if eA:
#         #     s2 = '-{}'.format(eA)
#         if sA and sA != '0':
#             s1 = '{}: +{}'.format(sysP[index], sA)
#         if eA and eA != '0':
#             s2 = '{}: -{}'.format(sysV[index], eA)
#         if s1 and s2:
#             s = s1 + ', ' + s2
#         else:
#             s = (s1 or s2) or s
#         s += '\n{}, {}, {}'.format(categ[index], pritok[index], vityajka[index]).replace(' ,', '').replace('\n, ', '\n')  # Доп инфа
#         if s[-1] == ' ':
#         	s = s[:-1]
#         if s[-1] == ',':
#         	s = s[:-1]
#         sp.LookupParameter('Комментарии').Set(s)
#         sp.LookupParameter('ADSK_Наименование приточной системы').Set(sysP[index])
#         sp.LookupParameter('ADSK_Наименование вытяжной системы').Set(sysV[index])
#         # sp.LookupParameter('Фильтр').Set(pritok[index])
#         # sp.LookupParameter('Вытяжка').Set(vityajka[index])
#         #if sA: s.LookupParameter('ADSK_Расчетный приток').Set(int(sA))
#         # elif not(s.LookupParameter('ADSK_Расчетный приток').AsDouble()): s.LookupParameter('ADSK_Расчетный приток').Set(sA)
#         #if eA: s.LookupParameter('ADSK_Расчетная вытяжка').Set(eA)
#         # elif not(s.LookupParameter('ADSK_Расчетная вытяжка').AsDouble()): s.LookupParameter('ADSK_Расчетная вытяжка').Set(eA)
#     else:
#         print('В таблице нет {} {}'.format(
#             n, sp.LookupParameter('Имя').AsString()))

# print('--------------------------------------------------------')

# for sp in sorted(list(set(nums))):
#     if sp not in existNums:
#         print('В модели нет {}'.format(sp))
# t.Commit()
