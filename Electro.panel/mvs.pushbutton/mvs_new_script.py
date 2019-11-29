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

    els = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().ToElements()

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
                len2 = len(doc.GetElement(el.GetTypeId()).LookupParameter('КабельСигнальный').AsString().split('/'))
                name = el.LookupParameter('Тип').AsValueString()
                if len1 > len2:
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
        cir.origin.LookupParameter('Имя нагрузки').Set(cir.els[0].LookupParameter('Тип').AsValueString())  # Используется имя только первого элемента цепи! Могут быть проблемы при разнородных шлейфах
        cir.origin.LookupParameter('Помещение').Set(cir.els[0].LookupParameter('Помещение').AsString())  # –//–
        cir.panel.LookupParameter('Имя панели').Set(cir.panel.LookupParameter('Тип').AsValueString())

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
        result = re.findall(r'^\(\(.+\)\)', wire_as_slashed_string)
        if result:
            wire_as_slashed_string = wire_as_slashed_string.split('/')
        i = 0
        for number, cir in sorted(cirs_grouped_by_element[el_id], key=lambda (n, c): n):
            cir.wire_sring_and_number_by_element = (wire_as_slashed_string, i)
            i += 1


    def get_wire_name(cir):
        wire_sring = cir.wire_sring_and_number_by_element[0]
        number_by_element = cir.wire_sring_and_number_by_element[1]
        # print('{} ------- {} ------- {} ------- {} ------- {}'.format(cir.id, cir.els[0].LookupParameter('Тип').AsValueString(), cir.panel.LookupParameter('Тип').AsValueString(), number_by_element, wire_sring))
        return wire_sring.split('/')[number_by_element]


    set_progress(pb, 40)

    nearest = Electrical.ElectricalCircuitPathMode.AllDevices
    for cir in Cir.objects:
        if get_wire_name(cir) == 'USB 2.0':
            cir.origin.CircuitPathMode = nearest  # наикратчайший путь
        else:
            value = doc.GetElement(cir.els[0].GetTypeId()).LookupParameter('КабельСиловой').AsString()
            if value:
                if 'пут' in value:
                    cir.CircuitPathMode = nearest  # наикратчайший путь

    t.Commit()
    t = Transaction(doc, 'Part 2')
    t.Start()

    default = 2000
    dict_ = {
       # КабельСигнальный                                         Наименование и техническая…                Тип, марка…                        Запас на разделку
        'FTP 4x2x0,52 cat 5e':                                  ['Кабель парной скрутки',                   'FTP 4x2x0,52 cat 5e',              default],  # noqa
        'FTP 5e':                                               ['Кабель парной скрутки',                   'FTP 2x2x0,52 cat 5e',              default],  # noqa
        'RG174':                                                ['Кабель коаксильный',                      'RG174',                            default],  # noqa
        'RG-174':                                               ['Кабель коаксильный',                      'RG174',                            default],  # noqa
        'RG-6U (75 Ом)':                                        ['Кабель коаксильный',                      'RG-6U (75 Ом)',                    default],  # noqa
        'SAT 703 B':                                            ['Кабель коаксильный',                      'SAT 703 B',                        default],  # noqa
        'UTP 4x2x0,52 cat 5e':                                  ['Кабель парной скрутки',                   'UTP 4x2x0,52 cat 5e',              default],  # noqa
        'UTP 4x2x0,52 cat 6a':                                  ['Кабель парной скрутки',                   'UTP 4x2x0,52 cat 6a',              default],  # noqa
        'UTP 4x2x0,52 cat 6e':                                  ['Кабель парной скрутки',                   'UTP 4x2x0,52 cat 6e',              default],  # noqa
        'Акустический кабель Audiocore Primary Wire M ACS0102': ['Кабель акустический',                     'Audiocore Primary Wire M ACS0102', default],  # noqa
        'ВВГнг(А) 3х0,75':                                      ['Кабель силовой',                          'ВВГнг(А) 3х0,75',                  default],  # noqa
        'ВВГнг(А)-LS 3x1,5':                                    ['Кабель силовой',                          'ВВГнг(А)-LS 3x1,5',                default],  # noqa
        'КВВГЭ 10х0,75':                                        ['Кабель',                                  'КВВГЭ 10х0,75',                    default],  # noqa
        'ШВВП 2х0,75':                                          ['Кабель силовой',                          'ШВВП 2х0,75',                      default],  # noqa
        '20 жильный управляющий кабель':                        ['Кабель (Уточнить)',                       '20 жил (Условно)',                 default],  # noqa
        '3xRCA':                                                ['Кабель передачи аудио-видеосигналов',     '3xRCA',                            default],  # noqa
        'HDMI(не включать)':                                    ['Кабель мультимедийный высокоскоростной',  'HDMI(не включать)',                200    ],  # noqa
        'HDMI':                                                 ['Кабель мультимедийный высокоскоростной',  'HDMI',                             200    ],  # noqa
        'USB 2.0(не включать)':                                 ['Кабель',                                  'USB 2.0(не включать)',             100    ],  # noqa
        'USB 2.0':                                              ['Кабель',                                  'USB 2.0',                          100    ],  # noqa
        'VGA':                                                  ['Кабель видеоинтерфейсный',                'VGA',                              100    ],  # noqa
        '2×Alarm':                                              ['Кабель специальный',                      '2×Alarm',                          0      ],  # noqa
        '4×Alarm':                                              ['Кабель специальный',                      '4×Alarm',                          0      ],  # noqa
    }

    set_progress(pb, 50)

    for cir in Cir.objects:
        wire_name = get_wire_name(cir)
        cir.origin.LookupParameter('КабельЦепи').Set(wire_name)


        name = dict_[wire_name][0] if wire_name in dict_ else 'Ошибка, добавь в скрипт ' + wire_name
        cir.origin.LookupParameter('Наименование').Set(name)

        coefficient_for_piece = 1.05
        coefficient_for_lenght = 1.5

        lenght = cir.origin.Length * k * coefficient_for_piece + dict_[wire_name][2]
        mark = dict_[wire_name][1] if wire_name in dict_ else 'Ошибка, добавь в скрипт ' + wire_name
        calc_method = 1
        if wire_name in ['HDMI', 'HDMI(не включать)', 'USB 2.0(не включать)', 'VGA']:
            for i in [1, 2, 3, 5, 20]:
                if lenght <= i * 1000:
                    lenght = i * 1000
                    break
            mark = mark + ', ' + str(lenght / 1000) + ' м'
        elif wire_name == 'USB 2.0':
            mark = mark + ', 3 м'
        else:
            calc_method = 0
            lenght = ceil(cir.origin.Length * k * coefficient_for_lenght + dict_[wire_name][2] / 1000)
        # print('{} {} {} '.format(lenght, ceil(lenght / 1000) * 1000, ceil(lenght) / k, ))
        cir.origin.LookupParameter('Длина с запасом').Set(ceil(lenght / 1000) * 1000 / k)
        cir.origin.LookupParameter('Тип, марка').Set(mark)
        cir.origin.LookupParameter('Способ расчёта').Set(calc_method)
        cir.origin.LookupParameter('Количество').Set(1 if calc_method else ceil(lenght / 1000))
        cir.origin.LookupParameter('Единицы измерения').Set('шт.' if calc_method else 'м')


    set_progress(pb, 60)

    els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().ToElements()
    # for el in filter(lambda x: 'Фейк' in x.LookupParameter('Тип').AsValueString(), els):
    #     doc.Delete(el.Id)
    [doc.Delete(i.Id) for i in els if 'Фейк' in i.LookupParameter('Тип').AsValueString()]
    els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().ToElements()
    # for el in filter(lambda x: 'Фейк' not in x.LookupParameter('Тип').AsValueString(), els):
    #     el.LookupParameter('Количество').Set(0)
    [i.LookupParameter('Количество').Set(0) for i in els if 'Фейк' not in i.LookupParameter('Тип').AsValueString()]
    symbols = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsElementType().ToElements()
    # symbols = list(filter(lambda x: 'Фейк' in x.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString(), symbols))
    symbols = [i for i in symbols if 'Фейк' in i.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()]
    for symbol in symbols:
        symbol.Activate() if not symbol.IsActive else None
        symbol.LookupParameter('Описание').Set('')
        symbol.LookupParameter('Комментарии к типоразмеру').Set('')
        symbol.LookupParameter('Ключевая пометка').Set('')
    level = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()[0]
    location = XYZ(0, 0, 0)
    done = []
    for cir in Cir.objects:
        mark = cir.origin.LookupParameter('Тип, марка').AsString()
        if mark not in done:
            done.append(mark)
            if len(done) > len(symbols):
                raise Exception('Добавь фейков')##################################################################################################### создать скриптом
            symbol = symbols[len(done) - 1]
            symbol.LookupParameter('Описание').Set(cir.origin.LookupParameter('Наименование').AsString())
            symbol.LookupParameter('Комментарии к типоразмеру').Set(mark)
            symbol.LookupParameter('Ключевая пометка').Set(cir.origin.LookupParameter('Единицы измерения').AsString())
            symbol.LookupParameter('Стоимость').Set(200)
        else:
            # symbol = list(filter(lambda x: x.LookupParameter('Комментарии к типоразмеру').AsString() == mark, symbols))[0]
            symbol = [i for i in symbols if i.LookupParameter('Комментарии к типоразмеру').AsString() == mark][0]
        el = doc.Create.NewFamilyInstance(location, symbol, level, Structure.StructuralType.NonStructural)
        # cir.LookupParameter('Способ расчёта').Set(sposobRaschyota[i])
        el.LookupParameter('Количество').Set(cir.origin.LookupParameter('Количество').AsDouble())
        el.LookupParameter('Цепь').Set(str(cir.id))
        el.LookupParameter('Помещение').Set(cir.origin.LookupParameter('Помещение').AsString())
        location += XYZ(0, -0.05, 0)

    set_progress(pb, 70)

    # symbols = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsElementType().ToElements()
    # symbols = list(filter(lambda x: 'офр' in x.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString(), symbols))
    symbols = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsElementType().ToElements()
    symbols = [i for i in symbols if 'офр' in i.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()]

    location = XYZ(-1, 0, 0)
    els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().ToElements()
    # for el in filter(lambda x: 'офр' in x.LookupParameter('Тип').AsValueString(), els):
    #     doc.Delete(el.Id)
    [doc.Delete(i.Id) for i in els if 'офр' in i.LookupParameter('Тип').AsValueString()]
    els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().ToElements()
    for symbol in symbols:
        if not symbol.IsActive:
            symbol.Activate()
        symbol.LookupParameter('Стоимость').Set(300)
        # for vega in filter(lambda x: 'Compact A3' in x.LookupParameter('Тип').AsValueString() or 'Блок сетевой управляющий ВЕГА' in x.LookupParameter('Тип').AsValueString(), els):
        for vega in [i for i in els if i.LookupParameter('Тип').AsValueString() in ['Compact A3', 'Блок сетевой управляющий ВЕГА']]:
            el = doc.Create.NewFamilyInstance(location, symbol, level, Structure.StructuralType.NonStructural)
            kol = symbol.LookupParameter('Группа модели').AsString()
            el.LookupParameter('Количество').Set(float(kol))  # Может быть ошибка, если Группа модели не прописана (её нужно прописывать вручную каждому типу крепежа).
            el.LookupParameter('Помещение').Set(vega.LookupParameter('Помещение').AsString())
            location += XYZ(0, -0.05, 0)

    set_progress(pb, 80)

    #---------------------------------Линии по трактории-----------------
    els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Lines).WhereElementIsNotElementType().ToElements()
    [doc.Delete(i.Id) for i in els if '!мвс ' in i.LineStyle.Name]
    lineStyles = [i for i in FilteredElementCollector(doc).OfClass(GraphicsStyle) if '!мвс ' in i.Name]
    cats = doc.Settings.Categories
    for cir in Cir.objects:
        path = list(cir.origin.GetCircuitPath())
        # offsets = [None] + path[:-1], path, path[1:] + [None]  # https://stackoverflow.com/a/56654140
        offsets = path, path[1:] + [None]  # https://stackoverflow.com/a/56654140
        for point, next_point in list(zip(*offsets)):
            if next_point:
                line = Line.CreateBound(point, next_point)  # https://forums.autodesk.com/t5/revit-api-forum/how-to-crete-3d-modelcurves-to-avoid-exception-curve-must-be-in/td-p/8355936
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
                    lineStyles.append(cats.NewSubcategory(cats.get_Item(BuiltInCategory.OST_Lines ), '!мвс ' + name))
                model_line.LineStyle = [i for i in lineStyles if i.Name == '!мвс ' + name][0]

    set_progress(pb, 90)

    els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().ToElements()

    for el in els:
        if el.GetSubComponentIds():
            room_name = el.LookupParameter('Помещение').AsString()
            for sub_el_id in el.GetSubComponentIds():
                doc.GetElement(sub_el_id).LookupParameter('Помещение').Set(room_name)

    mydict = {}
    for el in els:
        symbol = doc.GetElement(el.GetTypeId())
        if symbol.LookupParameter('Изображение типоразмера').AsElementId().ToString() == '-1':
            if symbol.LookupParameter('URL').AsString() != 'Без УГО':
                name = symbol.LookupParameter('Имя типа').AsString()
                if name not in mydict.keys():
                    mydict[name] = []
                mydict[name].append(el.Id)

    if mydict:
        if len(mydict.keys()) == 1:
            print('Данный типоразмер не имеет УГО. Следует добавить Изображение типоразмера (310×167.png) либо в параметре URL указать "Без УГО"')
        else:
            print('Данные типоразмеры не имеют УГО. Следует добавить Изображение типоразмера (310×167.png) либо в параметре URL указать "Без УГО"')
        for name in mydict.keys():
            print('{}'.format(output.linkify(mydict[name], name)))

    set_progress(pb, 100)

    t.Commit()
    tg.SetName('{} ({}, {:.1f} с)'.format(tg.GetName(), time.strftime('%H:%M', time.gmtime()), time.time() - startTime))
    tg.Assimilate()
