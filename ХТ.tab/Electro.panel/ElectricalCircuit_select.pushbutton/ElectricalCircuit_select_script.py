# -*- coding: utf-8 -*-
"""Добавляет к выбору цепи выбранного оборудования. Отфильтровывает из выбора лишние категории"""
__title__ = 'Выбрать цепи\nоборудования'
__author__ = 'SG'

import re
import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *
from Autodesk.Revit.DB import ElementId, PartUtils, ViewOrientation3D, XYZ, FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter
import sys
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
k = 304.8


def natural_sorted(list, key=lambda s: s):
    """
    Sort the list into natural alphanumeric order.
    """
    def get_alphanum_key_func(key):
        convert = lambda text: int(text) if text.isdigit() else text  # noqa
        return lambda s: [convert(c) for c in re.split('([0-9]+)', key(s))]
    sort_key = get_alphanum_key_func(key)
    return sorted(list, key=sort_key)


sel = [doc.GetElement(elid) for elid in uidoc.Selection.GetElementIds()]

# sel = filter(lambda x: x.LookupParameter('Категория').AsValueString() == 'Электрооборудование', sel)
sel = [el for el in sel if el.LookupParameter('Категория').AsValueString() == 'Электрооборудование']

# if sel and __forceddebugmode__:
if sel and not __shiftclick__:
    from pyrevit import script
    output = script.get_output()
    system_ids_to_select = []
    for el in sel:
        if el.MEPModel.ElectricalSystems:
            for system in el.MEPModel.ElectricalSystems:
                if system.Id not in system_ids_to_select:
                    system_ids_to_select.append(system.Id)
        elif 'ейк' in el.LookupParameter('Тип').AsValueString():
            target = el.LookupParameter('Цепь').AsString()
            system_ids_to_select.append(ElementId(int(target)))
    data = []
    system_ids_to_select = list(set(system_ids_to_select))
    # print(system_ids_to_select)
    # 1/0
    system_ids_to_select = natural_sorted(system_ids_to_select,
        key=lambda id: (doc.GetElement(id).LookupParameter('Позиция конца').AsString() if doc.GetElement(id).LookupParameter('Позиция конца') else ''))
    system_ids_to_select = natural_sorted(system_ids_to_select,
        key=lambda id: doc.GetElement(id).LookupParameter('Позиция начала').AsString() if doc.GetElement(id).LookupParameter('Позиция начала') else '')
    system_ids_to_select = natural_sorted(system_ids_to_select,
        key=lambda id: doc.GetElement(id).LookupParameter('Помещение цепи').AsString() if doc.GetElement(id).LookupParameter('Позиция начала') else '')
    for id in system_ids_to_select:
        el = doc.GetElement(id)
        data.append([
            output.linkify(id),
            el.LookupParameter('Помещение цепи').AsString() if el.LookupParameter('Помещение цепи') else el.LookupParameter('Помещение').AsString() if el.LookupParameter('Помещение') else '',
            # el.LookupParameter('Позиция начала').AsString() if el.LookupParameter('Позиция начала') else '',
            output.linkify(
                list(el.Elements)[0].Id,
                el.LookupParameter('Позиция начала').AsString() if el.LookupParameter('Позиция начала') else ''
            ),
            el.LookupParameter('Имя нагрузки').AsString() if el.LookupParameter('Имя нагрузки') else '',
            # el.LookupParameter('Позиция конца').AsString() if el.LookupParameter('Позиция конца') else '',
            output.linkify(
                # list(el.BaseEquipment)[0].Id,
                el.BaseEquipment.Id,
                el.LookupParameter('Позиция конца').AsString() if el.LookupParameter('Позиция конца') else ''
            ),
            el.LookupParameter('Панель').AsString(),
            el.LookupParameter('Тип, марка').AsString(),
            'Все устройства' if str(el.CircuitPathMode) == 'AllDevices' else\
                'Наиболее удалённое устройство' if str(el.CircuitPathMode) == 'FarthestDevice' else\
                'Пользоват.' if str(el.CircuitPathMode) == 'Custom' else el.CircuitPathMode,
            '{:.0f}'.format(el.PathOffset * k) if el.PathOffset else '',
            '{:.0f}'.format(el.LookupParameter('Длина').AsDouble() * k),
            '{:n}'.format(el.LookupParameter('Количество').AsDouble()),
            el.LookupParameter('Строгая длина, м').AsString() or '',
            # и почему в списке есть лишние цепи, например сервер # 2021.05.12: чта?
            ])
    # data = natural_sorted(data, key=lambda x: x[4])
    # data = natural_sorted(data, key=lambda x: x[2])
    #     print('{} Пом. {}: {} {} – {} {} {} Отступ={}'.format(
    #         output.linkify(id),
    #         el.LookupParameter('Помещение цепи').AsString() if el.LookupParameter('Помещение цепи') else el.LookupParameter('Помещение').AsString() if el.LookupParameter('Помещение') else '',
    #         el.LookupParameter('Позиция начала').AsString() if el.LookupParameter('Позиция начала') else '',
    #         el.LookupParameter('Имя нагрузки').AsString() if el.LookupParameter('Имя нагрузки') else '',
    #         el.LookupParameter('Позиция конца').AsString() if el.LookupParameter('Позиция конца') else '',
    #         el.LookupParameter('Панель').AsString(),
    #         el.LookupParameter('Тип, марка').AsString(),
    #         el.PathOffset * k,
    #     ))
    if len(system_ids_to_select) > 1:
        print('Выбрать все {}'.format(output.linkify(system_ids_to_select)))
    # data = [
    # ['row1', 'data', 'data', 80 ],
    # [output.linkify(ElementId(952627)), 'data', 'data', 45 ],
    # ]
    data = [[i + 1] + arr for i, arr in zip(range(len(data)), data)]
    output.print_table(
        table_data=data,
        columns=['№', 'Цепь', 'Помещение цепи', 'Позиция начала', 'Имя нагрузки', 'Позиция конца', 'Панель', 'Тип, марка', 'Режим траектории', 'Отступ', 'Длина revit', 'Количество, м/шт.', 'Строгая длина',],
        formats=['', '', '', ],
        # last_line_style='color:red;'
    )

# if sel and __shiftclick__:
#     if el.MEPModel.ElectricalSystems:
#         panel_ids = []
#         for system in el.MEPModel.ElectricalSystems:
#             if system.BaseEquipment and system.BaseEquipment.Id != el.Id:
#                 panel_ids.append(system.BaseEquipment.Id)
#         uidoc.Selection.SetElementIds(List[ElementId](panel_ids))

elif sel:
    sel_save = [el.Id for el in sel if 'ейк' not in el.LookupParameter('Тип').AsValueString()]

    system_ids_to_select = []
    for el in sel:
        # print(el.Id)
        if el.MEPModel.ElectricalSystems:
            for system in el.MEPModel.ElectricalSystems:
                if system.Id not in system_ids_to_select:
                    system_ids_to_select.append(system.Id)
        elif 'ейк' in el.LookupParameter('Тип').AsValueString():
            target = el.LookupParameter('Цепь').AsString()
            system_ids_to_select.append(ElementId(int(target)))



    uidoc.Selection.SetElementIds(List[ElementId](system_ids_to_select + sel_save))
