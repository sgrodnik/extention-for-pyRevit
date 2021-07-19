# -*- coding: utf-8 -*-
""""""
import time
startTime = time.time()

__title__ = 'Расчет\nMVS 2'
__author__ = 'SG'

import clr
clr.AddReference('System.Core')
import re
from System.Collections.Generic import *
from Autodesk.Revit.DB import ElementId, PartUtils, ViewOrientation3D, XYZ, FilteredElementCollector, BuiltInCategory, Transaction
from Autodesk.Revit.DB import TransactionGroup, BuiltInParameter, Line, Point, SketchPlane, Plane, Structure, Electrical, GraphicsStyle
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from math import ceil
from pyrevit import script
from pyrevit import forms
output = script.get_output()


def set_progress(pb, progress):
    pb.update_progress(progress, 100)
    pb.title = 'Расчет MVS: ' + str(progress) + ' %'


with forms.ProgressBar() as pb:

    set_progress(pb, 5)

    doc = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument
    k = 304.8

    default = 2000
    dict_ = {
       # КабельСигнальный                                        0 Наименование и техническая…              1 Тип, марка…                       2 Запас на разделку 3 Изготовитель
        'FTP 4x2x0,52 cat 5e':                                  ['Кабель парной скрутки',                   'FTP 4x2x0,52 cat 5e',              default,            'Россия'                                ],  # noqa
        'FTP 5e':                                               ['Кабель парной скрутки',                   'FTP 2x2x0,52 cat 5e',              default,            'Россия'                                ],  # noqa
        'RG174':                                                ['Кабель коаксильный',                      'RG174',                            default,            'Россия'                                ],  # noqa
        'RG-174':                                               ['Кабель коаксильный',                      'RG174',                            default,            'Россия'                                ],  # noqa
        'RG-6U (75 Ом)':                                        ['Кабель коаксильный',                      'RG-6U (75 Ом)',                    default,            'Россия'                                ],  # noqa
        'SAT 703 B':                                            ['Кабель коаксильный',                      'SAT 703 B',                        default,            'Россия'                                ],  # noqa
        'UTP 4x2x0,52 cat 5e':                                  ['Кабель парной скрутки',                   'UTP 4x2x0,52 cat 5e',              default,            'Россия'                                ],  # noqa
        'UTP 4x2x0,52 cat 6a':                                  ['Кабель парной скрутки',                   'UTP 4x2x0,52 cat 6',               default,            'Россия'                                ],  # noqa
        'UTP 4x2x0,52 cat 6e':                                  ['Кабель парной скрутки',                   'UTP 4x2x0,52 cat 6',               default,            'Россия'                                ],  # noqa
        'UTP 4x2x0,52 cat 6':                                   ['Кабель парной скрутки',                   'UTP 4x2x0,52 cat 6',               default,            'Россия'                                ],  # noqa
        'Акустический кабель Audiocore Primary Wire M ACS0102': ['Кабель акустический',                     'Audiocore Primary Wire M ACS0102', default,            'Россия'                                ],  # noqa
        'ACS0102':                                              ['Кабель акустический',                     'Audiocore Primary Wire M ACS0102', default,            'Россия'                                ],  # noqa
        'ВВГнг(А) 3х0,75':                                      ['Кабель силовой',                          'ВВГнг(А) 3х0,75',                  default,            'Россия'                                ],  # noqa
        'ВВГнг(А)-LS 3x1,5':                                    ['Кабель силовой',                          'ВВГнг(А)-LS 3x1,5',                default,            'Россия'                                ],  # noqa
        'КВВГЭ 10х0,75':                                        ['Кабель',                                  'КВВГЭ 10х0,75',                    default,            'Россия'                                ],  # noqa
        'ШВВП 2х0,75':                                          ['Кабель силовой',                          'КПСнг(А)-FRLS 1x2x0,75',           default,            'Россия'                                ],  # noqa
        'КПС':                                                  ['Кабель силовой',                          'КПСнг(А)-FRLS 1x2x0,75',           default,            'Россия'                                ],  # noqa
        '20 жильный управляющий кабель':                        ['Кабель (Уточнить)',                       '20 жил (Условно)',                 default,            'Россия'                                ],  # noqa
        '3xRCA':                                                ['Кабель передачи аудио-видеосигналов',     '3xRCA',                            default,            'Россия'                                ],  # noqa
        'HDMI(не включать)':                                    ['Кабель мультимедийный высокоскоростной',  'HDMI(не включать)',                200    ,            'Россия'                                ],  # noqa
        'HDMI':                                                 ['Кабель мультимедийный высокоскоростной',  'HDMI',                             200    ,            'Россия'                                ],  # noqa
        'USB 2.0(не включать)':                                 ['Кабель',                                  'USB 2.0(не включать)',             100    ,            'Россия'                                ],  # noqa
        'USB 2.0':                                              ['Кабель',                                  'USB 2.0',                          100    ,            'Россия'                                ],  # noqa
        'VGA':                                                  ['Кабель видеоинтерфейсный',                'VGA',                              100    ,            'Россия'                                ],  # noqa
        '2×Alarm':                                              ['Кабель специальный',                      '2×Alarm',                          0      ,            'Россия'                                ],  # noqa
        '4×Alarm':                                              ['Кабель специальный',                      '4×Alarm',                          0      ,            'Россия'                                ],  # noqa
        '---':                                                  ['------------------',                      '-------------',                    1000   ,            '------'                                ],  # noqa
        'sdi':                                                  ['Кабель коаксильный белый (RG 6)',         'MARS HD, Cu/Al/CuSn, 96%, 75 Ом',  1000   ,            'Rexant'                                ],  # noqa
        'sdi4':                                                 ['Кабель коаксильный 4 в 1 (для компактной прокладки)', 'Quad-link 3G-SDI',     1000   ,            'Kramer / AJA'                          ],  # noqa
        'utp':                                                  ['Кабель парной скрутки огнестойкий',       'UTP 4×2×0,52 cat 6 ZH нг(А)-HF',   1000   ,            'Rexant'                                ],  # noqa
        'utpp':                                                 ['Пaтч-корд F/UTP, категория 6, PVC серый', 'F/UTP cat 6 RJ45-RJ45',            400    ,            'Rexant'                                ],  # noqa
        # 'utp2':      не используем двупарку                   ['Кабель парной скрутки',                   'UTP 2PR 24AWG CAT5 305м нг(А)-HF', 1000   ,            'Rexant'                                ],  # noqa
        'ftp':                                                  ['Кабель парной скрутки огнестойкий',       'FTP 4×2×0,52 cat 5e ZH нг(А)-HF',  1000   ,            'Rexant'                                ],  # noqa
        'hdmi':                                                 ['Кабель мультимедийный высокоскоростной',  'HDMI',                             200    ,            'Россия'                                ],  # noqa
        'vga':                                                  ['Кабель видеоинтерфейсный',                'VGA',                              100    ,            'Россия'                                ],  # noqa
        'svideo':                                               ['Кабель видеоинтерфейсный',                'S-Video',                          100    ,            'Россия'                                ],  # noqa
        'pwr':                                                  ['Кабель парной скрутки огнестойкий',       'КПСнг(А)-FRLS 1×2×0,75',           1000   ,            'Rexant'                                ],  # noqa
        'audio':                                                ['Кабель акустический',                     'Primary Wire M ACS0102',           1000   ,            'Audiocore'                             ],  # noqa
        'aud':                                                  ['Кабель акустический экранированный',      'LCS-4 (PL-4704)',                  1000   ,            'Россия'                                ],  # noqa
        'usb':                                                  ['Кабель',                                  'USB 2.0',                          100    ,            'Россия'                                ],  # noqa
        '4alarm':                                               ['Кабель специальный',                      '4×Alarm',                          0      ,            'ООО "Медицинские системы визуализации"'],  # noqa
        '2alarm':                                               ['Кабель специальный',                      '2×Alarm',                          0      ,            'ООО "Медицинские системы визуализации"'],  # noqa
        '6alarm':                                               ['Кабель специальный',                      '6×Alarm',                          0      ,            'ООО "Медицинские системы визуализации"'],  # noqa
        'opt':                                                  ['Кабель волоконно-оптический',             'ОБВ-М нг(А)-HFLTx G.651 0М3 ММ',   1000      ,         'ООО "Инкаб"'],  # noqa
    }

    els = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().ToElements()

    depr_names = {
        'Акустический кабель Audiocore Primary Wire M ACS0102': 'ACS0102',
        'Акустический кабель': 'ACS0102',
    }
    depr_by_name = {}
    for el in els:
        symbol = doc.GetElement(el.GetTypeId())
        for name in depr_names:
            if symbol.LookupParameter('КабельСигнальный').AsString():
                if name in symbol.LookupParameter('КабельСигнальный').AsString():
                    if name not in depr_by_name:
                        depr_by_name[name] = set()
                    depr_by_name[name].add(symbol.Id)

    if depr_by_name:
        clr.AddReference('System.Windows.Forms')
        clr.AddReference('IronPython.Wpf')
        from pyrevit import script

        xamlfile = script.get_bundle_file('window.xaml')
        import wpf
        from System import Windows


        class MyWindow(Windows.Window):
            def __init__(self):
                wpf.LoadComponent(self, xamlfile)
                self.lb.Content = 'Обнаружены устаревшие способы указания кабеля:'
                info = ''
                for name in depr_by_name:
                    info += 'Устаревший кабель: "{}" следует заменить на новый: "{}":\n'.format(name, depr_names[name])
                    info += '\n'.join([
                        '{:<50}'.format(doc.GetElement(i).LookupParameter('Имя типа').AsString())
                        + ': ' + doc.GetElement(i).LookupParameter('КабельСигнальный').AsString()
                        for i in depr_by_name[name]
                    ])
                    info += '\n\n'
                self.sv.Content = info

            def yes(self, sender, args):
                self.Close()
                t = Transaction(doc, '(Замена названий кабелей)')
                t.Start()
                for el in els:
                    symbol = doc.GetElement(el.GetTypeId())
                    for name in depr_names:
                        signal_cable = symbol.LookupParameter('КабельСигнальный').AsString()
                        if signal_cable:
                            if name in signal_cable:
                                symbol.LookupParameter('КабельСигнальный').Set(signal_cable.replace(name, depr_names[name]))
                t.Commit()

            def no(self, sender, args):
                self.Close()


        MyWindow().ShowDialog()

    for el in els:  # Проверки наличия панелей и количества КабелейСигнальных
        if el.MEPModel.ElectricalSystems:
            arr = []
            for cir in list(el.MEPModel.ElectricalSystems):
                if not cir.BaseEquipment:
                    name = el.LookupParameter('Тип').AsValueString()
                    cir_number = cir.LookupParameter('Номер цепи').AsString()
                    print('{} цепь {} не имеет панели'.format(cir_number, output.linkify(el.Id, name)))
                    raise Exception('Не указана панель')
                if cir.BaseEquipment.Id != el.Id:
                    arr.append(cir)
            if arr:
                len1 = len(arr)
                cabs = doc.GetElement(el.GetTypeId()).LookupParameter('КабельСигнальный').AsString().split('/')
                len2 = len([i for i in cabs if i != ''])
                name = el.LookupParameter('Тип').AsValueString()
                if len1 > len2 and doc.GetElement(el.GetTypeId()).LookupParameter('Имя типа').AsString() != 'Трассировка':
                    print('Ошибка: {}: подключений: {} КабельСигнальный: {}'.format(output.linkify(el.Id, name), len1, len2))
                # elif len1 < len2:
                #     print('Предупреждение: {}: подключений: {} КабельСигнальный: {}'.format(output.linkify(el.Id, name), len1, len2))

    cirs = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_ElectricalCircuit).WhereElementIsNotElementType().ToElements()

    set_progress(pb, 10)


    class Cir():
        objects = []

        def __init__(self, origin):
            self.__class__.objects.append(self)
            self.origin = origin
            self.id = origin.Id
            self.panel = origin.BaseEquipment
            self.els = list(origin.Elements)

    tg = TransactionGroup(doc, 'Расчет MVS 2')
    tg.Start()
    t = Transaction(doc, 'Part 1')
    t.Start()

    for cir in cirs:
        cir = Cir(cir)
        cir.origin.LookupParameter('Имя нагрузки').Set(cir.els[0].LookupParameter('Тип').AsValueString().replace(' значок выкл', '').split('((')[0])  # Используется имя только первого элемента цепи! Могут быть проблемы при разнородных шлейфах
        cir.origin.LookupParameter('Помещение').Set(cir.els[0].LookupParameter('Помещение').AsString() or '')  # –//–
        cir.origin.LookupParameter('Помещение панели').Set(cir.panel.LookupParameter('Помещение').AsString() or '') if cir.origin.LookupParameter('Помещение панели') else None
        cir.origin.LookupParameter('Помещение цепи').Set(cir.els[0].LookupParameter('Помещение').AsString().replace('Стойка ', '') or '') if cir.origin.LookupParameter('Помещение цепи') else None
        cir.panel.LookupParameter('Имя панели').Set(cir.panel.LookupParameter('Тип').AsValueString().split('((')[0])
        cir.origin.LookupParameter('Комментарии').Set('Панель "' + cir.panel.LookupParameter('Тип').AsValueString() + '", id цепи ' + str(cir.id))
        if cir.origin.LookupParameter('Позиция начала'):
            param = cir.els[0].LookupParameter('ADSK_Обозначение')
            designation = param.AsString() if param and param.AsString() else ''
            cost = doc.GetElement(cir.els[0].GetTypeId()).LookupParameter('Стоимость').AsDouble()  # Используется имя только первого элемента цепи! Могут быть проблемы при разнородных шлейфах
            cir.origin.LookupParameter('Позиция начала').Set('{:n}{}'.format(cost, '.' + designation if designation else '').replace(',', '.'))
            param = cir.panel.LookupParameter('ADSK_Обозначение')
            designation = param.AsString() if param and param.AsString() else ''
            cost = doc.GetElement(cir.panel.GetTypeId()).LookupParameter('Стоимость').AsDouble()
            cir.origin.LookupParameter('Позиция конца').Set('{:n}{}'.format(cost, '.' + designation if designation else '').replace(',', '.'))

    set_progress(pb, 20)

    cirs_grouped_by_element = {}
    for cir in Cir.objects:
        element = cir.els[0]
        if element.Id not in cirs_grouped_by_element:
            cirs_grouped_by_element[element.Id] = []
        for connector in element.MEPModel.ConnectorManager.Connectors:
            all_refs = list(connector.AllRefs)
            if all_refs:
                if all_refs[0].Owner.Id == cir.id:
                    cirs_grouped_by_element[element.Id].append((connector.Id, cir))

    set_progress(pb, 30)

    for el_id in cirs_grouped_by_element:
        el = doc.GetElement(el_id)
        wire_as_slashed_string = doc.GetElement(el.GetTypeId()).LookupParameter('КабельСигнальный').AsString()
        brackets = re.findall(r'^\(\(.{1,5}\)\)', wire_as_slashed_string)
        if brackets:
            wires = wire_as_slashed_string.replace(brackets[0], '').split('/')
            newstr = ''
            for i in wires:
                newstr += i + ('' if '((' in i else brackets[0]) + '/'
            newstr += newstr[:-1]
            wire_as_slashed_string = newstr
        i = 0
        for number, cir in sorted(cirs_grouped_by_element[el_id], key=lambda (n, c): n):
            cir.wire_sring_and_number_by_element = (wire_as_slashed_string, i)
            i += 1

    def get_wire_name(cir):
        wire_sring = cir.wire_sring_and_number_by_element[0]
        number_by_element = cir.wire_sring_and_number_by_element[1]
        wire_name = wire_sring.split('/')[number_by_element]
        brackets = re.findall(r'\(\(.+\)\)', wire_name)
        if brackets:
            wire_name = wire_name.replace(brackets[0], '')
        return wire_name

    def get_strict_length(cir):
        param = cir.origin.LookupParameter('Строгая длина, м')
        instance_strict_length = param.AsString() if param and param.AsString() else ''
        addition = 0
        if instance_strict_length:
            if instance_strict_length.isdigit():
                instance_strict_length = int(instance_strict_length)
            elif instance_strict_length[0] == '+':                        # См. ниже
                addition = int(instance_strict_length[1:])
                instance_strict_length = None
            elif instance_strict_length[0] == '-':
                addition = -int(instance_strict_length[1:])
                instance_strict_length = None
        if not instance_strict_length:
            wire_sring = cir.wire_sring_and_number_by_element[0]
            number_by_element = cir.wire_sring_and_number_by_element[1]
            wire_name = wire_sring.split('/')[number_by_element]
            brackets = re.findall(r'\(\(.+\)\)', wire_name)
            type_strict_length = None
            if brackets:
                type_strict_length = float(brackets[0].replace('((', '').replace('))', '').replace(',', '.')) * 1000
            return type_strict_length         # + addition * 1000 2020.10.29 недоделал :( проблема в том, что тут возвращается строгая длина, а не прибавка. Надо лезть глубже.
        return instance_strict_length * 1000  # + addition * 1000 2020.10.29 недоделал :( проблема в том, что тут возвращается строгая длина, а не прибавка. Надо лезть глубже.

    set_progress(pb, 40)

    nearest = Electrical.ElectricalCircuitPathMode.AllDevices
    for cir in Cir.objects:
        if get_wire_name(cir) == 'USB 2.0':
            cir.origin.CircuitPathMode = nearest  # наикратчайший путь
        else:
            value = doc.GetElement(cir.els[0].GetTypeId()).LookupParameter('КабельСиловой').AsString()
            if value:
                if 'пут' in value:
                    cir.origin.CircuitPathMode = nearest  # наикратчайший путь

    t.Commit()
    t = Transaction(doc, 'Part 2')
    t.Start()

    set_progress(pb, 50)

    coefficient_for_piece = 1.05
    coefficient_for_lenght = 1.25
    piece_names_list = ['HDMI', 'HDMI(не включать)', 'USB 2.0(не включать)', 'VGA', 'hdmi', 'usb', 'vga', 'svideo', 'utpp']
    piece_lenght_list = [1, 2, 3, 5, 20]

    for cir in Cir.objects:
        wire_name = get_wire_name(cir)
        # cir.origin.LookupParameter('КабельЦепи').Set(wire_name)

        name = dict_[wire_name][0] if wire_name in dict_ else 'Ошибка, добавь в скрипт ' + wire_name
        # print(cir.origin.Id)
        cir.origin.LookupParameter('Наименование').Set(name)  #2021.02.05 странно, откуда тут взялся этот параметр

        reserve = dict_[wire_name][2] if wire_name in dict_ else 0
        strict_length = get_strict_length(cir)
        if strict_length:
            lenght = strict_length
        else:
            if cir.origin.LookupParameter('Кабель - Длина'):
                lenghtSE = cir.origin.LookupParameter('Кабель - Длина').AsDouble()
            else:
                lenghtSE = cir.origin.Length

            lenght = lenghtSE * k * coefficient_for_piece + reserve
        mark = dict_[wire_name][1] if wire_name in dict_ else 'Ошибка, добавь в скрипт ' + wire_name
        calc_method = 1
        if wire_name in piece_names_list:
            if strict_length:
                mark = mark + ', ' + str(int(lenght / 1000)) + ' м'
            else:
                for i in piece_lenght_list:
                    if lenght <= i * 1000:
                        lenght = i * 1000
                        break
                mark = mark + ', ' + str(lenght / 1000) + ' м'
        elif wire_name == 'USB 2.0':
            lenght = 2000 if not strict_length else strict_length
            mark = mark + ', ' + str(lenght / 1000) + ' м'
        else:
            calc_method = 0
            if strict_length:
                lenght = strict_length
            else:
                if cir.origin.LookupParameter('Кабель - Длина'): # 2021.05.27 вероятно это лишний повтор, см 25 строк выше
                    lenghtSE = cir.origin.LookupParameter('Кабель - Длина').AsDouble()
                else:
                    lenghtSE = cir.origin.Length
                lenght = ceil(lenghtSE * k * coefficient_for_lenght + reserve / 1000)
        cir.origin.LookupParameter('Длина с запасом').Set(ceil(lenght / 1000) * 1000 / k)
        cir.origin.LookupParameter('Тип, марка').Set(mark)
        cir.origin.LookupParameter('КабельЦепи').Set(mark)
        # cir.origin.LookupParameter('Способ расчёта').Set(calc_method)
        cir.origin.LookupParameter('Количество').Set(1 if calc_method else ceil(lenght / 1000))
        cir.origin.LookupParameter('Единицы измерения').Set('шт.' if calc_method else 'м')


    set_progress(pb, 60)

    els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().ToElements()
    [doc.Delete(i.Id) for i in els if 'Фейк' in i.LookupParameter('Тип').AsValueString()]
    els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().ToElements()
    [i.LookupParameter('Количество').Set(0) for i in els if 'Фейк' not in i.LookupParameter('Тип').AsValueString()]
    symbols = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsElementType().ToElements()
    symbols = [i for i in symbols if 'Фейк' in i.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()]
    for symbol in symbols:
        symbol.Activate() if not symbol.IsActive else None
        symbol.LookupParameter('Описание').Set('')
        symbol.LookupParameter('Комментарии к типоразмеру').Set('')
        symbol.LookupParameter('Ключевая пометка').Set('')
    level = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()[0]
    location = XYZ(0, 0, 0)
    done = []
    for cir in sorted(Cir.objects, key=lambda x: x.origin.LookupParameter('Наименование').AsString() + x.origin.LookupParameter('Тип, марка').AsString()):
        mark = cir.origin.LookupParameter('Тип, марка').AsString()
        if mark not in done:
            done.append(mark)
            if len(done) > len(symbols):
                # raise Exception('Добавь фейков')##################################################################################################### создать скриптом
                name = symbols[-1].Name.split()[0] + str(int(symbols[-1].Name.split()[-1]) + 1) if symbols[-1].Name.split()[-1].isdigit() else ' 1'
                symbols.append(symbols[-1].Duplicate(name))
            symbol = symbols[len(done) - 1]
            symbol.LookupParameter('Описание').Set(cir.origin.LookupParameter('Наименование').AsString())
            symbol.LookupParameter('Комментарии к типоразмеру').Set(mark)
            symbol.LookupParameter('Ключевая пометка').Set(cir.origin.LookupParameter('Единицы измерения').AsString())
            symbol.LookupParameter('Стоимость').Set(200 + len(done))
            # wire = doc.GetElement(cir.els[0].GetTypeId()).LookupParameter('КабельСигнальный').AsString()
            wire_name = get_wire_name(cir)
            # symbol.LookupParameter('Изготовитель').Set(dict_[wire_name][3])
            # symbol.LookupParameter('Изготовитель').Set(dict_[wire_name][3] if wire_name in dict_ else 'Ошибка, добавь в скрипт ' + wire_name)
            # print(wire_name)
            # print(type(wire_name))
            symbol.LookupParameter('Изготовитель').Set(dict_.get(wire_name, [0, 1, 2, 'Ошибка, добавь в скрипт '])[3])
        else:
            symbol = [i for i in symbols if i.LookupParameter('Комментарии к типоразмеру').AsString() == mark][0]
        el = doc.Create.NewFamilyInstance(location, symbol, level, Structure.StructuralType.NonStructural)
        el.LookupParameter('Количество').Set(cir.origin.LookupParameter('Количество').AsDouble())
        el.LookupParameter('Цепь').Set(str(cir.id))
        el.LookupParameter('ADSK_Группирование') and el.LookupParameter('ADSK_Группирование').Set('3. Кабельная продукция')
        el.LookupParameter('Помещение').Set(cir.origin.LookupParameter('Помещение').AsString())
        el.LookupParameter('Комментарии').Set(cir.origin.LookupParameter('Комментарии').AsString()) if cir.origin.LookupParameter('Комментарии').AsString() else None
        el.LookupParameter('Часть').Set(cir.origin.LookupParameter('Часть').AsString()) if cir.origin.LookupParameter('Часть') else None
        location += XYZ(0, -0.05, 0)

    set_progress(pb, 70)

    symbols = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsElementType().ToElements()
    symbols = [i for i in symbols if 'офр' in i.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()]

    location = XYZ(-1, 0, 0)
    els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().ToElements()
    [doc.Delete(i.Id) for i in els if 'офр' in i.LookupParameter('Тип').AsValueString()]
    els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().ToElements()
    for symbol in symbols:
        if not symbol.IsActive:
            symbol.Activate()
        symbol.LookupParameter('Стоимость').Set(300)
        for vega in [i for i in els if i.LookupParameter('Тип').AsValueString() in ['Compact A3', 'Блок сетевой управляющий ВЕГА']]:
            el = doc.Create.NewFamilyInstance(location, symbol, level, Structure.StructuralType.NonStructural)
            kol = symbol.LookupParameter('Группа модели').AsString().replace('Фейк', '')
            el.LookupParameter('Количество').Set(float(kol))  # Может быть ошибка, если Группа модели не прописана (её нужно прописывать вручную каждому типу крепежа).
            el.LookupParameter('Помещение').Set(vega.LookupParameter('Помещение').AsString() or '')
            location += XYZ(0, -0.05, 0)

    set_progress(pb, 80)

    #---------------------------------Линии по трактории-----------------
    els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Lines).WhereElementIsNotElementType().ToElements()
    [doc.Delete(i.Id) for i in els if '!мвс ' in i.LineStyle.Name]
    lineStyles = [i for i in FilteredElementCollector(doc).OfClass(GraphicsStyle) if '!мвс ' in i.Name]
    cats = doc.Settings.Categories
    errors = set()
    for cir in Cir.objects:
        path = list(cir.origin.GetCircuitPath())
        # offsets = [None] + path[:-1], path, path[1:] + [None]  # https://stackoverflow.com/a/56654140
        offsets = path, path[1:] + [None]  # https://stackoverflow.com/a/56654140
        for point, next_point in list(zip(*offsets)):
            if next_point:
                try:
                    line = Line.CreateBound(point, next_point)  # https://forums.autodesk.com/t5/revit-api-forum/how-to-crete-3d-modelcurves-to-avoid-exception-curve-must-be-in/td-p/8355936
                except Exception:
                    print('{} не удалось создать линии'.format(cir.id))
                    continue
                direction = line.Direction
                x, y, z = direction.X, direction.Y, direction.Z
                normal = XYZ(z - y, x - z, y - x)
                # normal = XYZ.BasisZ.CrossProduct(line.Direction)
                # if not normal.GetLength():
                #     normal = XYZ(1, 0, 0)
                plane = Plane.CreateByNormalAndOrigin(normal, point)
                sketchPlane = SketchPlane.Create(doc, plane)
                model_line = doc.Create.NewModelCurve(line, sketchPlane)
                name = get_wire_name(cir)
                if '!мвс ' + name not in [i.Name for i in lineStyles]:
                    # print(name)
                    # print(cir.els[0].Id)
                    lineStyles.append(cats.NewSubcategory(cats.get_Item(BuiltInCategory.OST_Lines), '!мвс ' + name))
                # model_line.LineStyle = [i for i in lineStyles if i.Name == '!мвс ' + name][0]
                for i in lineStyles:
                    if i.Name == '!мвс ' + name:
                        try:  # Срабатывает только со второго запуска скрипта. Не критично (Вероятно, можно исправить перезапуском подтранзакции)
                            model_line.LineStyle = i
                        except:
                            errors.add(name)
                        break
    if errors:
        for name in errors:
            print('Предупреждение: Не назначен стиль линии для "{}". Попробуйте запустить скрипт второй раз'.format(name))
        print('-' * 100)

    set_progress(pb, 90)

    els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().ToElements()

    for el in els:
        if el.GetSubComponentIds():
            room_name = el.LookupParameter('Помещение').AsString() or ''
            for sub_el_id in el.GetSubComponentIds():
                doc.GetElement(sub_el_id).LookupParameter('Помещение').Set(room_name)

    mydict = {}
    for el in els:
        symbol = doc.GetElement(el.GetTypeId())
        if symbol.LookupParameter('Изображение типоразмера').AsElementId().ToString() == '-1':
            # print(symbol.Id)
            if (not symbol.LookupParameter('URL').AsString()) or ('Без УГО' not in symbol.LookupParameter('URL').AsString()):
                name = symbol.LookupParameter('Имя типа').AsString()
                if name not in mydict.keys():
                    mydict[name] = []
                mydict[name].append(el.Id)
        symbol.LookupParameter('Стоимость').Set(0) if not symbol.LookupParameter('Стоимость').AsDouble() else None

    if mydict:
        if len(mydict.keys()) == 1:
            print('Данный типоразмер не имеет УГО. Следует добавить Изображение типоразмера (310×167.png) либо в параметре URL указать "Без УГО"')
        else:
            print('Данные типоразмеры не имеют УГО. Следует добавить Изображение типоразмера (310×167.png) либо в параметре URL указать "Без УГО"')
        print(' '.join(['{}'.format(output.linkify(mydict[name], name)) for name in mydict.keys()]))
        # for name in mydict.keys():
        #     print('{}'.format(output.linkify(mydict[name], name)))

    set_progress(pb, 98)

    els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ConduitFitting).WhereElementIsNotElementType().ToElements()
    for el in els:
        el.LookupParameter('Смещение_') and el.LookupParameter('Смещение_').Set(el.LookupParameter('Смещение').AsDouble())


    import json
    els = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_CableTray).WhereElementIsNotElementType().ToElements())
    els += list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Conduit).WhereElementIsNotElementType().ToElements())
    els += list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ConduitFitting).WhereElementIsNotElementType().ToElements())

    for el in els:
        symbol = doc.GetElement(el.GetTypeId())
        if not symbol.LookupParameter('Сортамент'):
            continue
        param = symbol.LookupParameter('Сортамент').AsString()
        settings = json.loads(param or '{}')
        if not settings:
            continue
        # if el.LookupParameter('Диаметр (Размер по каталогу)'):
        #     size = str(round(float(el.LookupParameter('Диаметр (Размер по каталогу)').AsDouble() * k), 1))
        if el.LookupParameter('Размер'):
            size = el.LookupParameter('Размер').AsString().split()[0].split('-')[0]
        # print(size)
        # print(settings)
        el.LookupParameter('ADSK_Наименование').Set(settings['ADSK_Наименование'][size])
        el.LookupParameter('ADSK_Марка').Set(settings['ADSK_Марка'][size])
        el.LookupParameter('ADSK_Код изделия').Set(settings['ADSK_Код изделия'][size])
        if el.LookupParameter('Длина'):
            length = el.LookupParameter('Длина').AsDouble() * k * float(settings['Множитель длины'])
            el.LookupParameter('ХТ Длина ОВ').Set(length)
        # if el.LookupParameter('Кабельная трасса'):  # смотри "Test: Кабельная трасса в Комментарии"
        #     res = el.LookupParameter('Кабельная трасса').AsString()
        #     res = ', '.join(res.split('\n'))
        #     el.LookupParameter('Комментарии').Set(res)

    set_progress(pb, 100)

    t.Commit()
    tg.SetName('{} ({}, {:.1f} с)'.format(tg.GetName(), time.strftime('%H:%M', time.localtime()), time.time() - startTime))
    tg.Assimilate()
