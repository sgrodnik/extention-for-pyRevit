# -*- coding: utf-8 -*-
"""Расчёт таблицы рассчета нагрузок"""
__title__ = 'Расчет\nТРН'
__author__ = 'SG'

import re
import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *

from Autodesk.Revit.DB import XYZ, FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

k = 304.8

def main():

    # sel = uidoc.Selection.GetElementIds()
    # els = [doc.GetElement(id) for id in sel]

    elCirs = FilteredElementCollector(doc)\
        .OfCategory(BuiltInCategory.OST_ElectricalCircuit)\
        .WhereElementIsNotElementType().ToElements()

    t = Transaction(doc, "Расчет ТРН")
    t.Start()

    for elCir in elCirs:
        s = ''
        for elem in list(elCir.Elements):
            if elem.LookupParameter('Наименование потребителя'):
                if elem.LookupParameter('Наименование потребителя').AsString():
                    naim = elem.LookupParameter('Наименование потребителя').AsString().encode('utf-8')
                else:
                    naim = elem.Name.encode('utf-8')
            else:
                naim = elem.Name.encode('utf-8')
            s += '{}; '.format(naim)
            # s += '{}; '.format(elem.Name.encode('utf-8'))
        # print(s)
        # print(elCir.LookupParameter('Имя нагрузки').Set(s.decode('utf-8')))
        # print(elCir.LookupParameter('Имя нагрузки').AsString().encode('utf-8'))

    global counter
    counter = 0
    hcirs = [cir for cir in elCirs if 'Щит' in cir.BaseEquipment.Name]
    global branches
    branches = []
    for cir in hcirs:
        preBranch = []
        find_longer_path(cir, preBranch)

    for branch in branches:
        name = branch[0].BaseEquipment.LookupParameter('Имя щита').AsString()
        for cir in branch:
            cir.LookupParameter('Имя щита').Set(name)

    panelsClusterBranches = {}
    for branch in branches:
        panel = branch[0].BaseEquipment.Id.IntegerValue
        if panel not in panelsClusterBranches:
            panelsClusterBranches[panel] = {}
        clusterHead = branch[0].Id.IntegerValue
        if clusterHead not in panelsClusterBranches[panel]:
            panelsClusterBranches[panel][clusterHead] = []
        panelsClusterBranches[panel][clusterHead].append(branch)

    pcb = panelsClusterBranches

    for panel in pcb.keys():
        for clusterHead in pcb[panel].keys():
            grNames = []
            for branch in pcb[panel][clusterHead]:
                for cir in branch:
                    for el in list(cir.Elements):
                        grNames.append(el.LookupParameter('Номер группы').AsString())
            grNames = filter(lambda x: x != '', natural_sorted(list(set(grNames))))
            if len(grNames) > 1:
                for name in grNames:
                    if '!' in name:
                        actualName = name
                        break
                    else:
                        actualName = grNames[0]
                actualName = actualName.replace('!', '')
                print('Внимание! Номера групп {} заменены на {}'.format(', '.join(grNames), actualName))
            else:
                actualName = grNames[0]
            for branch in pcb[panel][clusterHead]:
                for cir in branch:
                    cir.LookupParameter('Номер группы').Set(actualName)
    for elCir in elCirs:
        name = elCir.LookupParameter('Номер группы').AsString()
        for elem in list(elCir.Elements):
            elem.LookupParameter('Номер группы').Set(name)

    for panel in pcb.keys():
        for clusterHead in pcb[panel].keys():
            cableNames = []
            for branch in pcb[panel][clusterHead]:
                for cir in branch:
                    for el in list(cir.Elements):
                        if el.LookupParameter('Кабель'):
                            if el.LookupParameter('Кабель').AsString():
                                cableNames.append(el.LookupParameter('Кабель').AsString())
            cableNames = filter(lambda x: x != '', natural_sorted(list(set(cableNames))))
            if len(cableNames) == 0:
                cableNames = ['ВВГнг(А)-LSLTx']
            if len(cableNames) > 1:
                for name in cableNames:
                    if '!' in name:
                        actualName = name
                        break
                    else:
                        actualName = cableNames[0]
                actualName = actualName.replace('!', '')
                print('Внимание! В группе {} кабели {} заменены на {}'.format(pcb[panel][clusterHead][0][0].LookupParameter('Номер группы').AsString(), ', '.join(cableNames), actualName))
            else:
                actualName = cableNames[0]
            for branch in pcb[panel][clusterHead]:
                for cir in branch:
                    cir.LookupParameter('Кабель').Set(actualName)
    for elCir in elCirs:
        name = elCir.LookupParameter('Кабель').AsString()
        for elem in list(elCir.Elements):
            if elem.LookupParameter('Кабель'):
                elem.LookupParameter('Кабель').Set(name)

    panelsGroupsBranches = {}
    for branch in branches:
        panelName = branch[0].LookupParameter('Имя щита').AsString()
        if panelName not in panelsGroupsBranches:
            panelsGroupsBranches[panelName] = {}
        groupName = branch[0].LookupParameter('Номер группы').AsString()
        if groupName not in panelsGroupsBranches[panelName]:
            panelsGroupsBranches[panelName][groupName] = []
        panelsGroupsBranches[panelName][groupName].append(branch)


    for elCir in elCirs:
        # print(elCir)
        arr = list(elCir.Elements)
        # if arr[0].Name == 'Коробка распределительная, 220 В/220 В, Однофазная Тип системы, 3 Провода':
        #     elCir.LookupParameter('Номер группы').Set(arr[0].LookupParameter('Номер группы').AsString())
        # elif elCir.BaseEquipment.Name == 'Коробка распределительная, 220 В/220 В, Однофазная Тип системы, 3 Провода':
        #     elCir.LookupParameter('Номер группы').Set(elCir.BaseEquipment.LookupParameter('Номер группы').AsString())
        #     for i in arr:
        #         i.LookupParameter('Номер группы').Set(elCir.BaseEquipment.LookupParameter('Номер группы').AsString())
        # else:
        #     elCir.LookupParameter('Номер группы').Set(arr[0].LookupParameter('Номер группы').AsString())

        if len(arr) == 1:
            if arr[0].LookupParameter('Установленная мощность'):
                elCir.LookupParameter('Установленная мощность').Set(arr[0].LookupParameter('Установленная мощность').AsDouble())
            elif doc.GetElement(arr[0].GetTypeId()).LookupParameter('Установленная мощность'):
                elCir.LookupParameter('Установленная мощность').\
                    Set(doc.GetElement(arr[0].GetTypeId()).LookupParameter('Установленная мощность').AsDouble())
            else:
                elCir.LookupParameter('Установленная мощность').Set(0)
        elif len(arr) > 1:
            lst = []
            for i in arr:
                if i.LookupParameter('Установленная мощность'):
                    lst.append(i.LookupParameter('Установленная мощность').AsDouble())
                    # print('1')
                    # print(lst)
                else:
                    type = doc.GetElement(i.GetTypeId())
                    lst.append(type.LookupParameter('Установленная мощность').AsDouble())
                    # print('2')
                    # print(lst)
            elCir.LookupParameter('Установленная мощность').Set(sum(lst))

    pgb = panelsGroupsBranches

    for panel in pgb.keys():
        for group in pgb[panel].keys():
            power = 0
            for branch in pgb[panel][group]:
                for cir in branch:
                    power += cir.LookupParameter('Установленная мощность').AsDouble()
            for branch in pgb[panel][group]:
                for cir in branch:
                    cir.LookupParameter('Расчетная мощность').Set(power)

    for panel in pgb.keys():
        for group in pgb[panel].keys():
            condition = False
            for branch in pgb[panel][group]:
                for cir in branch:
                    for el in list(cir.Elements):
                        if el.LookupParameter('Установленная мощность'):
                            if el.LookupParameter('Установленная мощность').AsDouble() / 10763.9104167097 >= 3:
                                condition = True
                                break
                        elif doc.GetElement(el.GetTypeId()).LookupParameter('Установленная мощность'):
                            if doc.GetElement(el.GetTypeId()).LookupParameter('Установленная мощность').AsDouble() / 10763.9104167097 >= 3:
                                condition = True
                                break
                    if condition: break
            if condition: char = 'D'
            else: char = 'C'
            for branch in pgb[panel][group]:
                for cir in branch:
                    cir.LookupParameter('Характеристика аппарата защиты').Set(char)

    # for group in groups:
    #     condition = False
    #     for cir in groups[group]:
    #         for el in list(cir.Elements):
    #             if el.LookupParameter('Установленная мощность'):
    #                 if el.LookupParameter('Установленная мощность').AsDouble() / 10763.9104167097 >= 3:
    #                     condition = True
    #                     break
    #             elif doc.GetElement(el.GetTypeId()).LookupParameter('Установленная мощность'):
    #                 if doc.GetElement(el.GetTypeId()).LookupParameter('Установленная мощность').AsDouble() / 10763.9104167097 >= 3:
    #                     condition = True
    #                     break
    #         if condition: break
    #     if condition: char = 'D'
    #     else: char = 'C'
    #     for cir in groups[group]:
    #         cir.LookupParameter('Характеристика аппарата защиты').Set(char)

####################################
    # d = {}
    # groups = {}
    # for elCir in elCirs: # ХУЕТА! Надо считать для разных щитов.
    #     # print(d)
    #     if elCir.LookupParameter('Номер группы').AsString() not in d:
    #         d[elCir.LookupParameter('Номер группы').AsString()] = []
    #         groups[elCir.LookupParameter('Номер группы').AsString()] = []
    #     d[elCir.LookupParameter('Номер группы').AsString()].append(elCir.LookupParameter('Установленная мощность').AsDouble())
    #     groups[elCir.LookupParameter('Номер группы').AsString()].append(elCir)
    # sums = {}
    # for i in d:
    #     sums[i] = sum(d[i])
    # for elCir in elCirs:
    #     elCir.LookupParameter('Расчетная мощность').Set(sums[elCir.LookupParameter('Номер группы').AsString()])
##################################
    for elCir in elCirs:
        arr = list(elCir.Elements)
        if arr[0].Name == 'Коробка распределительная, 220 В/220 В, Однофазная Тип системы, 3 Провода':
            ks = float(arr[0].LookupParameter('Коэффициент спроса').AsDouble())
            elCir.LookupParameter('Коэффициент спроса').Set(ks if ks > 0 else 1)
        elif elCir.BaseEquipment.Name == 'Коробка распределительная, 220 В/220 В, Однофазная Тип системы, 3 Провода':
            ks = float(elCir.BaseEquipment.LookupParameter('Коэффициент спроса').AsDouble())
            elCir.LookupParameter('Коэффициент спроса').Set(ks if ks > 0 else 1)
            for i in arr:
                if i.LookupParameter('Коэффициент спроса'):
                    i.LookupParameter('Коэффициент спроса').Set(ks if ks > 0 else 1)
        else:
            # print('-------------------')
            # print(elCir.BaseEquipment.Id)
            if elCir.BaseEquipment.LookupParameter('Коэффициент спроса'):
                ks = float(elCir.BaseEquipment.LookupParameter('Коэффициент спроса').AsDouble())
                elCir.LookupParameter('Коэффициент спроса').Set(ks if ks > 0 else 1)
            elif arr[0].LookupParameter('Коэффициент спроса'):
                ks = float(arr[0].LookupParameter('Коэффициент спроса').AsDouble())
                elCir.LookupParameter('Коэффициент спроса').Set(ks if ks > 0 else 1)

    for elCir in elCirs:
        arr = list(elCir.Elements)
        if arr[0].Name == 'Коробка распределительная, 220 В/220 В, Однофазная Тип системы, 3 Провода':
            ks = float(arr[0].LookupParameter('Коэффициент спроса расчетный').AsDouble())
            elCir.LookupParameter('Коэффициент спроса расчетный').Set(ks if ks > 0 else 1)
        elif elCir.BaseEquipment.Name == 'Коробка распределительная, 220 В/220 В, Однофазная Тип системы, 3 Провода':
            ks = float(elCir.BaseEquipment.LookupParameter('Коэффициент спроса расчетный').AsDouble())
            elCir.LookupParameter('Коэффициент спроса расчетный').Set(ks if ks > 0 else 1)
            for i in arr:
                if i.LookupParameter('Коэффициент спроса расчетный'):
                    i.LookupParameter('Коэффициент спроса расчетный').Set(ks if ks > 0 else 1)
        else:
            # print('-------------------')
            # print(elCir.BaseEquipment.Id)
            if elCir.BaseEquipment.LookupParameter('Коэффициент спроса расчетный'):
                ks = float(elCir.BaseEquipment.LookupParameter('Коэффициент спроса расчетный').AsDouble())
                elCir.LookupParameter('Коэффициент спроса расчетный').Set(ks if ks > 0 else 1)
            elif arr[0].LookupParameter('Коэффициент спроса расчетный'):
                ks = float(arr[0].LookupParameter('Коэффициент спроса расчетный').AsDouble())
                elCir.LookupParameter('Коэффициент спроса расчетный').Set(ks if ks > 0 else 1)

    for elCir in elCirs:
        arr = list(elCir.Elements)
        if arr[0].Name == 'Коробка распределительная, 220 В/220 В, Однофазная Тип системы, 3 Провода':
            ks = float(arr[0].LookupParameter('Cos φ').AsDouble())
            elCir.LookupParameter('Cos φ').Set(ks if ks > 0 else 0.85)
        elif elCir.BaseEquipment.Name == 'Коробка распределительная, 220 В/220 В, Однофазная Тип системы, 3 Провода':
            ks = float(elCir.BaseEquipment.LookupParameter('Cos φ').AsDouble())
            elCir.LookupParameter('Cos φ').Set(ks if ks > 0 else 0.85)
            # for i in arr:
            #     if i.LookupParameter('Cos φ'):
            #         i.LookupParameter('Cos φ').Set(ks if ks > 0 else 0.85)
            type = doc.GetElement(i.GetTypeId())
            if arr[0].LookupParameter('Cos φ') and arr[0].LookupParameter('Cos φ').AsDouble() > 0:
                ks = float(arr[0].LookupParameter('Cos φ').AsDouble())
                elCir.LookupParameter('Cos φ').Set(ks if ks > 0 else 0.85)
            if type.LookupParameter('Cos φ'):
                ks = float(type.LookupParameter('Cos φ').AsDouble())
                elCir.LookupParameter('Cos φ').Set(ks if ks > 0 else 0.85)
        else:
            # print('-------------------')
            # print(elCir.BaseEquipment.Id)
            if elCir.BaseEquipment.LookupParameter('Cos φ'):
                ks = float(elCir.BaseEquipment.LookupParameter('Cos φ').AsDouble())
                elCir.LookupParameter('Cos φ').Set(ks if ks > 0 else 0.85)
            elif arr[0].LookupParameter('Cos φ'):
                ks = float(arr[0].LookupParameter('Cos φ').AsDouble())
                elCir.LookupParameter('Cos φ').Set(ks if ks > 0 else 0.85)

    for branch in branches:
        for i, cir in enumerate(reversed(branch)):
            if i == 0: cir.LookupParameter('Установленная мощность ветки').Set(cir.LookupParameter('Установленная мощность').AsDouble())
            else: cir.LookupParameter('Установленная мощность ветки').Set(0)
            if list(branch[-1].Elements)[0].LookupParameter('Напряжение цепи').AsDouble() == 380: u = 380
            else: u = 220
            cir.LookupParameter('Напряжение цепи').Set(u)

    for branch in branches:
        for i, cir in enumerate(reversed(branch)):
            if 'Щит' in cir.BaseEquipment.Name:
                break
            superCir = list(reversed(branch))[i+1]
            superCir.LookupParameter('Установленная мощность ветки').Set(
                superCir.LookupParameter('Установленная мощность ветки').AsDouble()
                + cir.LookupParameter('Установленная мощность ветки').AsDouble())

    for elCir in elCirs:
        P = elCir.LookupParameter('Установленная мощность ветки').AsDouble() / 10763.9104167097
        U = elCir.LookupParameter('Напряжение цепи').AsDouble()
        cos = elCir.LookupParameter('Cos φ').AsDouble()
        if U == 380:
            try:
                I = P / (U / 1000 * cos * 1.7320508075688772)
            except:
                continue
        else:
            try:
                I = P / (U / 1000 * cos)
            except:
                continue
        elCir.LookupParameter('Ток на участке').Set(I)

    for elCir in elCirs:
        P = elCir.LookupParameter('Расчетная мощность').AsDouble() / 10763.9104167097
        U = elCir.LookupParameter('Напряжение цепи').AsDouble()
        cos = elCir.LookupParameter('Cos φ').AsDouble()
        if U == 380:
            try:
                I = P / (U / 1000 * cos * 1.7320508075688772)
            except:
                continue
        else:
            try:
                I = P / (U / 1000 * cos)
            except:
                continue
        elCir.LookupParameter('Ток').Set(I)

    noms = {10: 1.5,
            16: 2.5,
            20: 4,
            25: 4,
            32: 6,
            40: 6,
            50: 10,
            63: 16,
            80: 25,
            100: 25,
            125: 35,
            160: 50,
            200: 70,
            225: 95,
            250: 120,
            315: 150,
            400: 240}

    for elCir in elCirs:
        for i in range(len(sorted(noms.keys()))):
            if sorted(noms.keys())[i] >= elCir.LookupParameter('Ток').AsDouble():
                elCir.LookupParameter('Номинал аппарата защиты').Set(sorted(noms.keys())[i])
                break
        a = elCir.LookupParameter('Номинал аппарата защиты').AsDouble()
        elCir.LookupParameter('Сечение кабеля').Set(noms[a])

    for elCir in elCirs:
        P = elCir.LookupParameter('Установленная мощность ветки').AsDouble() / 10763.9104167097
        L = elCir.LookupParameter('Длина').AsDouble() * 304.8 / 1000
        S = elCir.LookupParameter('Сечение кабеля').AsDouble()
        C = 12 if elCir.LookupParameter('Напряжение цепи').AsDouble() == 220 else 72
        dU = P * L / (S * C)
        elCir.LookupParameter('Падение').Set(dU)

    panelsCirs = {}
    # q = 0
    for elCir in elCirs:
        # print('#', q)
        # q += 1
        name = elCir.LookupParameter('Имя щита').AsString()
        if name in panelsCirs:
            panelsCirs[name].append(elCir)
        else:
            panelsCirs[name] = [elCir]
        # for i in panelsCirs[name]: print(i.LookupParameter('Ток').AsDouble())

    # print('---------------------------------------------')

    for panel in panelsCirs.keys():
        panelsCirs[panel].sort(key=lambda x: x.LookupParameter('Ток').AsDouble(), reverse=True)
        # for i in panelsCirs[panel]: print(i.LookupParameter('Ток').AsDouble())

    for sortedCirs in panelsCirs.keys():
        a, b, c = 0, 0, 0
        for cir in panelsCirs[sortedCirs]:
            if 'Щит' in cir.BaseEquipment.Name:
                if cir.LookupParameter('Напряжение цепи').AsDouble() == 380:
                    a += cir.LookupParameter('Ток').AsDouble()
                    b += cir.LookupParameter('Ток').AsDouble()
                    c += cir.LookupParameter('Ток').AsDouble()
                    cir.LookupParameter('Номер фазы').Set('1, 2, 3')
                elif a == min(0, 0, 0):
                    a += cir.LookupParameter('Ток').AsDouble()
                    cir.LookupParameter('Номер фазы').Set('1')
                elif b == min(0, 0, 0):
                    b += cir.LookupParameter('Ток').AsDouble()
                    cir.LookupParameter('Номер фазы').Set('2')
                else:
                    c += cir.LookupParameter('Ток').AsDouble()
                    cir.LookupParameter('Номер фазы').Set('3')
        phaseShift = (max(a, b, c) - min(a, b, c)) / max(a, b, c) * 100
        for cir in panelsCirs[sortedCirs]:
            cir.LookupParameter('Перекос фаз').Set(phaseShift)
            cir.LookupParameter('Ток по фазам').Set('{:.2f}, {:.2f}, {:.2f}'.format(float(a), float(b), float(c)).replace('.', ','))
        # print(a, b, c)
        # print(phaseShift)

    for branch in branches:
        phase = branch[0].LookupParameter('Номер фазы').AsString()
        for cir in branch[1:]:
            cir.LookupParameter('Номер фазы').Set(phase)

    # panelsBranches = {}
    for branch in branches:
        phase = branch[0].LookupParameter('Номер фазы').AsString()
        for cir in branch[1:]:
            cir.LookupParameter('Номер фазы').Set(phase)

    for panel in pgb.keys(): ####################################### сделать ненулевую нумерацию
        i = 0
        for group in natural_sorted(pgb[panel].keys()):
            i += 1
            for branch in pgb[panel][group]:
                for cir in branch:
                    cir.LookupParameter('Номер аппарата защиты').Set(i)

    for panel in pgb.keys():
        for group in natural_sorted(pgb[panel].keys()):
            brlenlst = []
            n = 0
            for branch in pgb[panel][group]:
                brlenlst.append(0)
                for i, cir in enumerate(branch):
                    if i != len(branch) - 1 or len(branch) == 1:
                        brlenlst[n] += cir.LookupParameter('Длина').AsDouble() * 304.8
                    else:
                        n += 1
            for branch in pgb[panel][group]:
                for cir in branch:
                    cir.LookupParameter('Длина расчетная').Set(max(brlenlst)) # Максимальная длина ветки не включая последний участок

    for panel in pgb.keys():
        for group in natural_sorted(pgb[panel].keys()):
            voltageDrop = 0
            for branch in pgb[panel][group]:
                for cir in branch:
                    vd = cir.LookupParameter('Падение').AsDouble()
                    if vd > voltageDrop:
                        voltageDrop = vd
            if voltageDrop > 5:
                gr = pgb[panel][group][0][0].LookupParameter('Номер группы').AsString()
                print('Внимание! Падение напряжения {} составляет {:.2f}'.format(gr, voltageDrop))
            for branch in pgb[panel][group]:
                for cir in branch:
                    cir.LookupParameter('Падение группы').Set(voltageDrop)

    for panel in pgb.keys():
        for group in pgb[panel].keys():
            cirIds = set()
            for branch in pgb[panel][group]:
                for cir in branch:
                    cirIds.add(str(cir.Id.IntegerValue))
            for branch in pgb[panel][group]:
                branch[0].LookupParameter('Цепи').Set(' '.join(cirIds))


    draft(panelsCirs)

    t.Commit()

    return branches


def natural_sorted(list, key=lambda s:s):
    """
    Sort the list into natural alphanumeric order.
    """
    def get_alphanum_key_func(key):
        convert = lambda text: int(text) if text.isdigit() else text
        return lambda s: [convert(c) for c in re.split('([0-9]+)', key(s))]
    sort_key = get_alphanum_key_func(key)
    return sorted(list, key=sort_key)


def find_longer_path(elCir, branch):
    global branches
    branch.append(elCir)
    end = list(list(elCir.Elements)[0].MEPModel.ElectricalSystems)
    end = [i for i in end if str(i.Id.IntegerValue) != str(elCir.Id.IntegerValue)]
    if len(end) == 0:
        branches.append(map(lambda x: x, branch))
        # print(map(lambda x: x.LookupParameter('Имя нагрузки').AsString(), branch))
        branch.pop()
        return 1
    for cir in end:
        find_longer_path(cir, branch)
    branch.pop()


def draft(panels):
    els = FilteredElementCollector(doc)\
        .OfCategory(BuiltInCategory.OST_GenericAnnotation)\
        .WhereElementIsNotElementType()\
        .ToElements()

    for el in els:
        if el.Name == 'Цепь' or el.Name == 'Щит':
            doc.Delete(el.Id)

    anns = FilteredElementCollector(doc)\
        .OfCategory(BuiltInCategory.OST_GenericAnnotation)\
        .WhereElementIsElementType()\
        .ToElements()

    cirFamily = list(filter(lambda x: x.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                                       .AsString() == 'Цепь', anns))[0]
    panelFamily = list(filter(lambda x: x.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                                         .AsString() == 'Щит', anns))[0]
    activeView = doc.ActiveView

    for y, panel in enumerate(natural_sorted(panels.keys())):
        el = doc.Create.NewFamilyInstance(XYZ(0, -y * 250 / k, 0), panelFamily, activeView)
        el.LookupParameter('qwe').Set(panel)
        for x, cir in enumerate(filter(lambda x: 'Щит' in x.BaseEquipment.Name, natural_sorted(panels[panel], key = lambda x: x.LookupParameter('Номер группы').AsString()))):
            el = doc.Create.NewFamilyInstance(XYZ(x * 16 / k, -y * 250 / k, 0), cirFamily, activeView)
            el.LookupParameter('Номер группы').Set(cir.LookupParameter('Номер группы').AsString())
            el.LookupParameter('Расчетная мощность').Set(round(cir.LookupParameter('Расчетная мощность').AsDouble() / 10763.9104167097, 2))
            el.LookupParameter('Cos φ').Set(round(cir.LookupParameter('Cos φ').AsDouble(), 2))
            el.LookupParameter('Ток').Set(round(cir.LookupParameter('Ток').AsDouble(), 2))
            el.LookupParameter('Падение группы').Set(round(cir.LookupParameter('Падение группы').AsDouble(), 2))
            el.LookupParameter('Длина расчетная').Set(round(cir.LookupParameter('Длина расчетная').AsDouble() / 1000, 0))
            el.LookupParameter('Номинал аппарата защиты').Set(round(cir.LookupParameter('Номинал аппарата защиты').AsDouble(), 0))
            el.LookupParameter('Характеристика аппарата защиты').Set(cir.LookupParameter('Характеристика аппарата защиты').AsString())
            el.LookupParameter('Имя нагрузки').Set(cir.LookupParameter('Имя нагрузки').AsString())
            el.LookupParameter('Цепи').Set(cir.LookupParameter('Цепи').AsString())
            sech = '{:.1f}'.format(cir.LookupParameter('Сечение кабеля').AsDouble()).replace('.0', '')
            n = 3 if cir.LookupParameter('Напряжение цепи').AsDouble() == 220 else 5
            name = cir.LookupParameter('Кабель').AsString()
            el.LookupParameter('Кабель').Set('{} {}x{}'.format(name, n, sech.replace('.', ',')))


with open('C:\\Users\\sgrodnik\\Desktop\\TRN.txt', 'w', buffering=1) as o:
    # sys.stdout = o
    # print("Start")
    OUT = main()
