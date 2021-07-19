# -*- coding: utf-8 -*-
"""Shift + Click откроет файл"""
__title__ = 'Заполнение пространств, ТБВ и ХОВС'
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


# path = doc.PathName.split(' ОВ v')[0] + ' ТВБ.txt'
# try:
#     src = open(path, 'r').read().decode("utf-8")[:-1]
# except IOError:
#     info = path + '\nЭтот файл должен содержать таблицу (разделители - табуляция).\nПервая строка должна содержать заголовки.\nПоследняя строка должна быть пустой.\n'
#     with open(path, 'w') as file:
#         file.write(info.encode("utf-8"))
#     os.startfile(path)
# else:
#     if __shiftclick__:
#         os.startfile(path)
#         sys.exit()


path = os.path.join(os.path.dirname(doc.PathName), 'ТВБ, ХОВС.txt')
file_started = False
try:
    src = open(path, 'r').read().decode("utf-8")
    if src:
        src = src[:-1] if src[-1] == '\n' else src
    if not src or src.startswith('\nЭтот'):
        raise IOError
except IOError:
    info = '\nЭтот файл должен содержать таблицу (разделители - табуляция).\nПервая строка должна содержать заголовки. Последняя строка может быть пустой.\nМежду таблицами одна пустая строка\n' + path
    with open(path, 'w') as file:
        file.write(info.encode("utf-8"))
    os.startfile(path)
    src = '\n\n'
    file_started = True

if __shiftclick__:
    if not file_started:
        os.startfile(path)
    sys.exit()

src_spaces, src_hovs = src.split('\n\n')
# src_filters = src_filters.split('\n')[1:]
# src_views = src_views.split('\n')[1:]

t = Transaction(doc, 'Заполнение пространств')
t.Start()

if src_spaces:
    titles = src_spaces.split('\n')[0].split('\t')
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
    _cleanlinessClass = titles.index('Класс чистоты') if 'Класс чистоты' in titles else titles.index('Класс чистоты по СанПиН 2.1.3.2630-10')
    _temperature = titles.index('Расчетная температура воздуха в помещении')
    _area = titles.index('Площадь, м²')
    _height = titles.index('Высота, м')
    _volume = titles.index('Объем, м³')
    _multiplicityInflow = titles.index('Кратность притока, N/ч')
    _multiplicityOutflow = titles.index('Кратность вытяжки, N/ч')
    _note = titles.index('Примечание')
    _outflow = titles.index('Переток') if 'Переток' in titles else None
    _heat = titles.index('ΣQ Суммарные теплопоступления') if 'ΣQ Суммарные теплопоступления' in titles else None


    raws = [x.split('\t') for x in src_spaces.split('\n')][1:]                 # raws
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
    outflow = [x[_outflow] for x in raws] if _outflow else []
    heat = [x[_heat] for x in raws] if _heat else []


    spaces = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElements()

    teplo = False

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
            if teplo:
                s = 'Q={}кВт'.format(float(heat[index]) / 1000).replace('.', ',')
            else:
                if sA and sA != '0':
                    s1 = '{}: +{}'.format(supplySystem[index], sA)
                if eA and eA != '0':
                    s2 = '{}: -{}'.format(exhaustSystem[index], eA)
            if s1 and s2:
                s = s1 + ', ' + s2
            else:
                s = (s1 or s2) or s
            if not teplo:
                s += '\n{}, {}, {}'.format('Класс ' + cleanlinessClass[index] if cleanlinessClass[index] else '', supplyDevices[index], exhaustDevices[index]).replace(' ,', '').replace('\n, ', '\n')
            if s[-1] == ' ':
                s = s[:-1]
            if s[-1] == ',':
                s = s[:-1]
            current_outflow = int(outflow[index]) if outflow else 0
            if current_outflow:
                space.LookupParameter('ADSK_Количество') and space.LookupParameter('ADSK_Количество').Set(abs(current_outflow))
                # s += '\nПереток: ' + str(current_outflow) + ' м³/ч'
            old_s = space.LookupParameter('Комментарии').AsString()
            old_s = old_s if old_s else ''
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

    # for index, el in enumerate(els):
    #     el.LookupParameter('№').Set(number[index] if number[index] else '‎')
    #     el.LookupParameter('Наименование помещения').Set(name[index])
    #     if name[index] == 'Итого ':
    #         el.LookupParameter('№').Set(number[index] if number[index] else '‎‎‎‎‎‎')
    #         el.LookupParameter('Наименование помещения').Set('‎' + ' ' * 115 + 'Итого')
    #     el.LookupParameter('Класс чистоты').Set(cleanlinessClass[index])
    #     el.LookupParameter('Расчетная температура воздуха в помещении').Set(temperature[index])
    #     el.LookupParameter('Площадь, м²').Set(area[index])
    #     el.LookupParameter('Высота, м').Set(height[index])
    #     el.LookupParameter('Объем, м³').Set(volume[index])
    #     el.LookupParameter('Кратность притока, N/ч').Set(multiplicityInflow[index])
    #     el.LookupParameter('Кратность вытяжки, N/ч').Set(multiplicityOutflow[index])
    #     el.LookupParameter('Расход приточного воздуха, м³/ч').Set(supplyAirFlow[index])
    #     el.LookupParameter('Расход удаляемого воздуха, м³/ч').Set(exhaustAirFlow[index])
    #     el.LookupParameter('Приточные  воздухораспределительные устройства').Set(supplyDevices[index])
    #     el.LookupParameter('Вытяжные  воздухораспределительные устройства').Set(exhaustDevices[index])
    #     el.LookupParameter('Приточная система').Set(supplySystem[index])
    #     el.LookupParameter('Вытяжная система').Set(exhaustSystem[index])
    #     el.LookupParameter('Примечание.').Set(note[index])

    for index, el in enumerate(els):
        el.LookupParameter('Кл 1').Set(number[index] if number[index] else '‎')
        el.LookupParameter('Кл 2').Set(name[index])
        if name[index] == 'Итого ':
            el.LookupParameter('Кл 1').Set(number[index] if number[index] else '‎‎‎‎‎‎')
            el.LookupParameter('Кл 2').Set('‎' + ' ' * 115 + 'Итого')
        if 'Итого по группе помещений' in name[index]:
            el.LookupParameter('Кл 1').Set(number[index] if number[index] else '‎‎‎‎‎‎')
            # el.LookupParameter('Кл 2').Set('‎' + ' ' * 77 + 'Итого по группе помещений')
            el.LookupParameter('Кл 2').Set('‎' + ' ' * 60 + name[index])
        el.LookupParameter('Кл 3').Set(cleanlinessClass[index])
        el.LookupParameter('Кл 4').Set(temperature[index])
        el.LookupParameter('Кл 5').Set(area[index])
        el.LookupParameter('Кл 6').Set(height[index])
        el.LookupParameter('Кл 7').Set(volume[index])
        el.LookupParameter('Кл 8').Set(multiplicityInflow[index] if multiplicityInflow[index] != '0' else '')
        el.LookupParameter('Кл 9').Set(multiplicityOutflow[index] if multiplicityOutflow[index] != '0' else '')
        el.LookupParameter('Кл 10').Set(supplyAirFlow[index] if supplyAirFlow[index] != '0' else '')
        el.LookupParameter('Кл 11').Set(exhaustAirFlow[index] if exhaustAirFlow[index] != '0' else '')
        el.LookupParameter('Кл 12').Set(supplyDevices[index])
        el.LookupParameter('Кл 13').Set(exhaustDevices[index])
        el.LookupParameter('Кл 14').Set(supplySystem[index])
        el.LookupParameter('Кл 15').Set(exhaustSystem[index])
        el.LookupParameter('Кл 16').Set(note[index])

if src_hovs:
    src_hovs = src_hovs.split('\n')
    # src_filters.pop(-1) if src_filters[-1].split('\t')[0] == '---' else None
    # src_h_ = [i.split('\t')[0] for i in src_filters]
    # src_h_ = [i.split('\t')[1] for i in src_filters]
    # src_h_ = [i.split('\t')[2] for i in src_filters]
    # src_h_ = [i.split('\t')[3] for i in src_filters]
    # src_h_ = [i.split('\t')[4] for i in src_filters]

    schedules = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Schedules).WhereElementIsNotElementType().ToElements()
    schedule = list(filter(lambda x: x.Name == 'Характеристика систем', schedules))[0]
    data = schedule.GetTableData().GetSectionData(SectionType.Body)

    for i in range(data.LastRowNumber):
        data.RemoveRow(i + 1)

    for _ in range(len(src_hovs)):
        data.InsertRow(0)

    els = FilteredElementCollector(doc, schedule.Id).ToElements()

    for index, el in enumerate(els):
        el.LookupParameter('Ключевая1').Set(src_hovs[index].split('\t')[0])
        # print(type(src_hovs[index].split('\t')))
        # print(len(src_hovs[index].split('\t')))
        names = src_hovs[index].split('\t')[1]
        if names == 'Итого':
            names = '‎' + ' ' * 128 + 'Итого'
        el.LookupParameter('Ключевая2').Set(names)
        el.LookupParameter('Ключевая3').Set(src_hovs[index].split('\t')[2])
        el.LookupParameter('Ключевая4').Set(src_hovs[index].split('\t')[3])
        el.LookupParameter('Ключевая5').Set(src_hovs[index].split('\t')[4])
        el.LookupParameter('Ключевая6').Set(src_hovs[index].split('\t')[5])
        el.LookupParameter('Ключевая7').Set(src_hovs[index].split('\t')[6])
        el.LookupParameter('Ключевая8').Set(src_hovs[index].split('\t')[7])
        el.LookupParameter('Ключевая9').Set(src_hovs[index].split('\t')[8])
        el.LookupParameter('Ключевая10').Set(src_hovs[index].split('\t')[9])
        el.LookupParameter('Ключевая11').Set(src_hovs[index].split('\t')[10])
        el.LookupParameter('Ключевая12').Set(src_hovs[index].split('\t')[11])
        el.LookupParameter('Ключевая13').Set(src_hovs[index].split('\t')[12])
        el.LookupParameter('Ключевая14').Set(src_hovs[index].split('\t')[13])
        el.LookupParameter('Ключевая15').Set(src_hovs[index].split('\t')[14])
        el.LookupParameter('Ключевая16').Set(src_hovs[index].split('\t')[15])
        el.LookupParameter('Ключевая17').Set(src_hovs[index].split('\t')[16])
        el.LookupParameter('Ключевая18').Set(src_hovs[index].split('\t')[17])
        el.LookupParameter('Ключевая19').Set(src_hovs[index].split('\t')[18])
        el.LookupParameter('Ключевая20').Set(src_hovs[index].split('\t')[19])
        el.LookupParameter('Ключевая21').Set(src_hovs[index].split('\t')[20])
        el.LookupParameter('Ключевая22').Set(src_hovs[index].split('\t')[21])
        el.LookupParameter('Ключевая23').Set(src_hovs[index].split('\t')[22])
        el.LookupParameter('Ключевая24').Set(src_hovs[index].split('\t')[23])
        el.LookupParameter('Ключевая25').Set(src_hovs[index].split('\t')[24])
        el.LookupParameter('Ключевая26').Set(src_hovs[index].split('\t')[25])
        el.LookupParameter('Ключевая27').Set(src_hovs[index].split('\t')[26])
        el.LookupParameter('Ключевая28').Set(src_hovs[index].split('\t')[27])

t.Commit()
