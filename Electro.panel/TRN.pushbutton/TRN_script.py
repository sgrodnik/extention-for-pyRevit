# -*- coding: utf-8 -*-
"""Расчёт таблицы рассчета нагрузок"""
__title__ = 'Расчет\nТРН'
__author__ = 'SG'

import time
startTime = time.time()
import re
import math
import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *

from Autodesk.Revit.DB import XYZ, FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

k = 304.8


def find_longer_path(elCir, branch):
    global branches
    branch.append(elCir)
    if 'Щит' in list(elCir.Elements)[0].Name:
        branches.append(map(lambda x: x, branch))
        return 1
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


def natural_sorted(list, key=lambda s: s):
    """
    Sort the list into natural alphanumeric order.
    """
    def get_alphanum_key_func(key):
        convert = lambda text: int(text) if text.isdigit() else text
        return lambda s: [convert(c) for c in re.split('([0-9]+)', key(s))]
    sort_key = get_alphanum_key_func(key)
    return sorted(list, key=sort_key)

fuse = 0
def find_depth(cir, counter):
    global fuse
    fuse += 1
    if fuse > 100:
        raise Exception('Ошибка: превышена глубина')
    panel = cir.BaseEquipment
    if not panel:
        return
    if 'Щит' in panel.Name:
        headCir = None
        for c in list(panel.MEPModel.ElectricalSystems):
            if list(c.Elements)[0].Id == panel.Id:
                headCir = c
                break
        if headCir:
            counter.append(headCir.Id.IntegerValue)
            find_depth(headCir, counter)
        else:
            return


t = Transaction(doc, 'name')
t.Start()

elCirs = FilteredElementCollector(doc)\
    .OfCategory(BuiltInCategory.OST_ElectricalCircuit)\
    .WhereElementIsNotElementType().ToElements()

UnpluggedFakeCirs = list(filter(lambda x: not x.BaseEquipment and 'Щит' in list(x.Elements)[0].Name, elCirs))
elCirs = list(filter(lambda x: x.BaseEquipment, elCirs))

global branches
branches = []
for cir in [cir for cir in elCirs if 'Щит' in cir.BaseEquipment.Name]:
    preBranch = []
    find_longer_path(cir, preBranch)

pCirs = {}  # panelsCircuits
for branch in branches:
    panel = branch[0].BaseEquipment.Id
    if panel not in pCirs:
        pCirs[panel] = []
    pCirs[panel].extend(branch)

for i in pCirs.keys():
    pCirs[i] = list(set(pCirs[i]))

for cir in filter(lambda x: 'Щит' in x.BaseEquipment.Name, elCirs):
    cnr = []
    find_depth(cir, cnr)
    cir.BaseEquipment.LookupParameter('Глубина').Set(len(cnr))

deeps = {}
for panel in pCirs.keys():
    depth = doc.GetElement(panel).LookupParameter('Глубина').AsDouble()
    for cir in pCirs[panel]:
        if depth not in deeps:
            deeps[depth] = []
        deeps[depth].append(cir)

# deeps[-1] = UnpluggedFakeCirs

###   ###   ###   ###   ###   ###   ###   ###   ###   ###   ###   ###   ###
   ###   ###   ###   ###   ###   ###   ###   ###   ###   ###   ###   ###   ###

for depth in sorted(deeps.keys(), reverse=True):
    elCirs = deeps[depth]

    global branches
    branches = []
    for cir in [cir for cir in elCirs if 'Щит' in cir.BaseEquipment.Name]:
        preBranch = []
        find_longer_path(cir, preBranch)

    pcb = {}  # panelsClusterBranches
    pcc = {}  # panelsClusterCircuits
    pCirs = {}  # panelsCircuits
    for branch in branches:
        panel = branch[0].BaseEquipment.Id
        clusterHead = branch[0]

        if panel not in pcb:
            pcb[panel] = {}
        if clusterHead not in pcb[panel]:
            pcb[panel][clusterHead] = []
        pcb[panel][clusterHead].append(branch)

        if panel not in pcc:
            pcc[panel] = {}
        if clusterHead not in pcc[panel]:
            pcc[panel][clusterHead] = []
        pcc[panel][clusterHead].extend(branch)

        if panel not in pCirs:
            pCirs[panel] = []
        pCirs[panel].extend(branch)

    for i in pCirs.keys():
        pCirs[i] = list(set(pCirs[i]))

    # for cir in elCirs:
    #     if cir.BaseEquipment.Id not in deeps:
    #         deeps[cir.BaseEquipment.Id] = []
    #     deeps[cir.BaseEquipment.Id].append(cir)

    childCirs = {cir.Id.IntegerValue: [] for cir in elCirs}
    for panel in pcc.keys():
        for clusterHead in pcc[panel].keys():
            cluster = pcb[panel][clusterHead]
            for branch in cluster:
                for cir in branch:
                    cir.LookupParameter('Цепи').Set('')
            for branch in cluster:
                reversedBranch = list(reversed(branch))
                for i, cir in enumerate(reversedBranch):
                    cirIds = [cir.Id.IntegerValue]
                    if childCirs[cir.Id.IntegerValue]:
                        cirIds.extend(childCirs[cir.Id.IntegerValue])
                    childCirs[cir.Id.IntegerValue].extend(set(cirIds))
                    childCirs[cir.Id.IntegerValue] = list(set(childCirs[cir.Id.IntegerValue]))
                    if 'Щит' in cir.BaseEquipment.Name:
                        break
                    superCir = reversedBranch[i+1]
                    superCirIds = [superCir.Id.IntegerValue]
                    if childCirs[superCir.Id.IntegerValue]:
                        superCirIds.extend(childCirs[superCir.Id.IntegerValue])
                    childCirs[superCir.Id.IntegerValue].extend(set(cirIds + superCirIds))
                    childCirs[superCir.Id.IntegerValue] = list(set(childCirs[superCir.Id.IntegerValue]))

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

    for panel in sorted(pcc.keys(), key=lambda x: doc.GetElement(x).LookupParameter('Имя щита').AsString()):
        for clusterHead in pcc[panel].keys():
            groupsCirs = pcc[panel][clusterHead]
            branches = pcb[panel][clusterHead]

            groupNums = []
            cableNames = []
            cApps = []
            totalPower = 0
            characteristic = 'C'
            kss = []
            ksrs = []
            coss = []
            voltages = []
            voltageDrop = 0
            groupConsumerSpaces = []
            groupConsumerNames = []

            panelName = groupsCirs[0].BaseEquipment.LookupParameter('Имя щита').AsString()

            ############################################################################### Из цепи
            for cir in groupsCirs:
                groupNums.append(cir.LookupParameter('Номер группы').AsString())
                cableNames.append(cir.LookupParameter('Кабель').AsString())
                cApps.append(cir.LookupParameter('Ком. аппарат').AsString())
                kss.append(cir.LookupParameter('Коэффициент спроса').AsDouble())
                ksrs.append(cir.LookupParameter('Коэффициент спроса расчетный').AsDouble())
                coss.append(cir.LookupParameter('Cos φ').AsDouble())
                voltages.append(cir.LookupParameter('Напряжение цепи').AsDouble())

                consumerNames = []
                consumerSpaces = []

                installedPower = 0

                branchInstalledPower = 0
                for idInt in childCirs[cir.Id.IntegerValue]:
                    for el in list(doc.GetElement(ElementId(idInt)).Elements):
                        param = el.LookupParameter('Установленная мощность')
                        param0 = doc.GetElement(el.GetTypeId()).LookupParameter('Установленная мощность')
                        branchInstalledPower += param.AsDouble() if param else param0.AsDouble() if param0 else 0

                ############################################################################### Из элементов цепи
                for el in list(cir.Elements):
                    if 'Щит' not in el.Name:
                        param = el.LookupParameter('Номер группы')
                        groupNums.append(param.AsString() if param else '-')

                    param = el.LookupParameter('Наименование потребителя')
                    naim = param.AsString() if param else el.Name
                    naim = naim if naim else el.Name
                    naim = 'Розеточная сеть' if 'озетка' in naim else naim
                    naim = 'Розетка заземления' if 'Розетка заземления' in el.Name else naim
                    naim = '(Распредкоробка)' if 'оробка' in naim else naim
                    naim = naim.replace(', 220 В/220 В, Однофазная Тип системы, 3 Провода', '')
                    consumerNames.append(naim)

                    space = el.Space[doc.GetElement(el.CreatedPhaseId)]
                    consumerSpaces.append(space.Number if space else '-')

                    groupConsumerSpaces.extend(consumerSpaces)
                    groupConsumerNames.extend(consumerNames)

                    param = el.LookupParameter('Кабель')
                    cableNames.append(param.AsString() if param else '')

                    param = el.LookupParameter('Ком. аппарат')
                    # 'qf': 'Автоматический выключатель',
                    # 'QF': 'Автоматический выключатель 2P',
                    # 'QS': 'Выключатель',
                    # 'QFD': 'Дифф',
                    # 'KM': 'Контактор',
                    # 'FU': 'Плавкая вставка',
                    # 'QSU': 'Рубильник с плавкой вставкой',
                    # 'QD': 'УЗО'
                    cApps.append(param.AsString() if param else 'QF' if 'IT сеть' in el.Name else 'qf' if 'LED' in el.Name else '')

                    param = el.LookupParameter('Установленная мощность')
                    param0 = doc.GetElement(el.GetTypeId()).LookupParameter('Установленная мощность')
                    installedPower += param.AsDouble() if param else param0.AsDouble() if param0 else 0

                    if installedPower / 10763.9104167097 >= 3: characteristic = 'D'

                    param = el.LookupParameter('Коэффициент спроса')
                    param0 = doc.GetElement(el.GetTypeId()).LookupParameter('Коэффициент спроса')
                    kss.append(param.AsDouble() if param else param0.AsDouble() if param0 else 0)

                    param = el.LookupParameter('Коэффициент спроса расчетный')
                    param0 = doc.GetElement(el.GetTypeId()).LookupParameter('Коэффициент спроса расчетный')
                    ksrs.append(param.AsDouble() if param else param0.AsDouble() if param0 else 0)

                    param = el.LookupParameter('Cos φ')
                    param0 = doc.GetElement(el.GetTypeId()).LookupParameter('Cos φ')
                    coss.append(param.AsDouble() if param else param0.AsDouble() if param0 else 0)

                    if 'Щит' not in el.Name:
                        param = el.LookupParameter('Напряжение цепи')
                        voltages.append(param.AsDouble() if param else 0)

                ############################################################################### Из цепи, но после суммирования
                totalPower += installedPower

                ############################################################################### В цепь (для цепи)
                cir.LookupParameter('Имя нагрузки').Set('; '.join(set(consumerNames)))

                cir.LookupParameter('Помещение нагрузки').Set('; '.join(natural_sorted(set(consumerSpaces))))

                cir.LookupParameter('Имя щита').Set(panelName if panelName else '-')

                cir.LookupParameter('Установленная мощность').Set(installedPower)

                cir.LookupParameter('Установленная мощность ветки').Set(branchInstalledPower)

            branchLens = []
            n = 0
            for branch in branches:
                branchLens.append(0)
                for i, cir in enumerate(branch):
                    if i != len(branch) - 1 or len(branch) == 1:
                        branchLens[n] += cir.LookupParameter('Длина').AsDouble() * 304.8
                    else:
                        n += 1

            groupConsumerSpaces = filter(lambda x: '-' not in x, groupConsumerSpaces)
            groupConsumerNames = filter(lambda x: 'оробка' not in x, groupConsumerNames)


            ############################################################################### Расчеты
            groupNums = natural_sorted(filter(lambda x: x and x != '‎', list(set(groupNums))))
            if len(groupNums) == 0:
                groupNum = '-'
            elif len(groupNums) > 1:
                for name in groupNums:
                    if '!' in name:
                        groupNum = name.replace('!', '')
                        break
                    else:
                        groupNum = groupNums[0]
                print('{}: Номера групп {} заменены на {}'.format(panelName, ', '.join(groupNums), groupNum))
            else:
                groupNum = groupNums[0].replace('!', '')

            cableNames = natural_sorted(filter(lambda x: x, list(set(cableNames))))
            if len(cableNames) == 0:
                if 'IT' in groupNum or 'A' in groupNum or 'А' in groupNum:
                    cableName = 'ВВГнг(А)-FRLSLTx'
                else:
                    cableName = 'ВВГнг(А)-LSLTx'
            elif len(cableNames) > 1:
                for name in cableNames:
                    if '!' in name:
                        cableName = name
                        break
                    else:
                        cableName = cableNames[0].replace('!', '')
                print('{}: В группе {} кабели {} заменены на {}'.format(panelName, groupNum, ', '.join(cableNames), cableName))
            else:
                cableName = cableNames[0].replace('!', '')

            cApps = natural_sorted(filter(lambda x: x, list(set(cApps))))
            # if len(cApps) == 0:
                # cApp = 'QFD'
            if len(cApps) > 1:
                for name in cApps:
                    if '!' in name:
                        cApp = name
                        break
                    else:
                        cApp = cApps[0].replace('!', '')
                print('{}: В группе {} аппараты защиты {} заменены на {}'.format(panelName, groupNum, ', '.join(cApps), cApp))
            else:
                cApp = cApps[0].replace('!', '') if cApps else 'QFD'

            kss = filter(lambda x: x, kss)
            tmp = {}
            if len(set(kss)) == 1:
                ks = kss[0]
            elif kss:
                for i in kss:
                    if i not in tmp:
                        tmp[i] = 0
                    else:
                        tmp[i] += 1
                ks = sorted(tmp.keys(), key=lambda x: tmp[x])[0]
                print('{}: В группе {} коэффициенты спроса {} заменены на {:.2f}'.format(panelName, groupNum, ', '.join(map(lambda x: '{:.2f}'.format(x), set(kss))), ks))
            else:
                ks = 1

            ksrs = filter(lambda x: x, ksrs)
            tmp = {}
            if len(set(ksrs)) == 1:
                ksr = ksrs[0]
            elif ksrs:
                for i in ksrs:
                    if i not in tmp:
                        tmp[i] = 0
                    else:
                        tmp[i] += 1
                ksr = sorted(tmp.keys(), key=lambda x: tmp[x])[0]
                print('{}: В группе {} коэффициенты спроса расчетные {} заменены на {:.2f}'.format(panelName, groupNum, ', '.join(map(lambda x: '{:.2f}'.format(x), set(ksrs))), ksr))
            else:
                ksr = 1

            coss = filter(lambda x: x, coss)
            tmp = {}
            if len(set(coss)) == 1:
                cos = coss[0]
            elif coss:
                for i in coss:
                    if i not in tmp:
                        tmp[i] = 0
                    else:
                        tmp[i] += 1
                cos = sorted(tmp.keys(), key=lambda x: tmp[x])[0]
                print('{}: В группе {} Cos φ {} заменены на {:.2f}'.format(panelName, groupNum, ', '.join(map(lambda x: '{:.2f}'.format(x), set(coss))), cos))
            else:
                cos = 0.85

            voltages = filter(lambda x: x, voltages)
            tmp = {}
            if len(set(voltages)) == 1:
                voltage = voltages[0]
            elif voltages:
                for i in voltages:
                    if i not in tmp:
                        tmp[i] = 0
                    else:
                        tmp[i] += 1
                voltage = sorted(tmp.keys(), key=lambda x: tmp[x])[0]
                print('{}: В группе {} напряжения цепи {} заменены на {:.2f}'.format(panelName, groupNum, ', '.join(map(lambda x: '{:.2f}'.format(x), set(voltages))), voltage))
            else:
                voltage = 220.0

            ############################################################################### В цепь (для группы)
            for cir in groupsCirs:
                cir.LookupParameter('Номер группы').Set(groupNum)
                cir.LookupParameter('Кабель').Set(cableName)
                cir.LookupParameter('Ком. аппарат').Set(cApp)
                cir.LookupParameter('Суммарная мощность группы').Set(totalPower)
                cir.LookupParameter('Характеристика аппарата защиты').Set(characteristic)
                cir.LookupParameter('Коэффициент спроса').Set(ks)
                cir.LookupParameter('Коэффициент спроса расчетный').Set(ksr)
                cir.LookupParameter('Cos φ').Set(cos)
                cir.LookupParameter('Напряжение цепи').Set(voltage)

                # это все тут, а не выше, так как зависит от напряжения, которое посчитано только сейчас
                P = cir.LookupParameter('Установленная мощность ветки').AsDouble() / 10763.9104167097
                I = P / (voltage / 1000 * cos * 1.7320508075688772) if voltage == 380 else P / (voltage / 1000 * cos)
                cir.LookupParameter('Ток на участке').Set(I)

                P = totalPower / 10763.9104167097
                I = P / (voltage / 1000 * cos * 1.7320508075688772) if voltage == 380 else P / (voltage / 1000 * cos)
                cir.LookupParameter('Ток').Set(I)

                for i in range(len(noms)):
                    if sorted(noms.keys())[i] >= I:
                        nom = sorted(noms.keys())[i]
                        cir.LookupParameter('Номинал аппарата защиты').Set(nom)
                        break

                cir.LookupParameter('Сечение кабеля').Set(noms[nom])

                P = cir.LookupParameter('Установленная мощность ветки').AsDouble() / 10763.9104167097
                L = cir.LookupParameter('Длина').AsDouble() * 304.8 / 1000
                S = noms[nom]
                C = 12 if voltage == 220 else 72
                dU = P * L / (S * C)
                cir.LookupParameter('Падение').Set(dU)

                if dU > voltageDrop:
                    voltageDrop = dU
                if voltageDrop > 5:
                    print('{}: В группе {} Падение напряжения составляет {:.2f}'.format(panelName, groupNum, voltageDrop))

                cir.LookupParameter('Длина расчетная').Set(max(branchLens)) # Максимальная длина ветки не включая последний участок
                cir.LookupParameter('Падение группы').Set(voltageDrop)
                cir.LookupParameter('Цепи').Set(' '.join(map(str, childCirs[cir.Id.IntegerValue])))
                cir.LookupParameter('Помещения группы').Set('; '.join(natural_sorted(set(groupConsumerSpaces))))
                cir.LookupParameter('Имя нагрузки группы').Set('; '.join(natural_sorted(set(groupConsumerNames))))

                ksr = cir.LookupParameter('Коэффициент спроса расчетный').AsDouble()
                kodn = 1
                ratedPower = totalPower / 10763.9104167097 * ksr * kodn
                # if 'Щит' not in list(cir.Elements)[0].Name:  # Странно, почему потребовалась эта проверка
                cir.LookupParameter('Расчетная мощность').Set(ratedPower)
                tg = math.tan(math.acos(cos))
                cir.LookupParameter('tg φ').Set(tg)
                Q = ratedPower * tg
                cir.LookupParameter('Q, квар').Set(Q)

                ############################################################################### В элементы цепи
                for el in list(cir.Elements):
                    el.LookupParameter('Номер группы').Set(groupNum)
                    el.LookupParameter('Кабель').Set(cableName) if el.LookupParameter('Кабель') else 1
                    el.LookupParameter('Ком. аппарат').Set(cApp) if el.LookupParameter('Ком. аппарат') else 1
                    el.LookupParameter('Коэффициент спроса').Set(ks) if el.LookupParameter('Коэффициент спроса') else 1
                    el.LookupParameter('Коэффициент спроса расчетный').Set(ksr) if el.LookupParameter('Коэффициент спроса расчетный') else 1
                    el.LookupParameter('Cos φ').Set(cos) if el.LookupParameter('Cos φ') else 1
                    el.LookupParameter('Напряжение цепи').Set(voltage) if el.LookupParameter('Напряжение цепи') else 1

    ##########################################################################################
    ##########################################################################################
    for i in pCirs.keys():
        a, b, c = 0, 0, 0
        for cir in sorted(pCirs[i], key=lambda x: x.LookupParameter('Ток').AsDouble(), reverse=True):
            if 'Щит' in cir.BaseEquipment.Name:
                if cir.LookupParameter('Напряжение цепи').AsDouble() == 380:
                    a += cir.LookupParameter('Ток').AsDouble()
                    b += cir.LookupParameter('Ток').AsDouble()
                    c += cir.LookupParameter('Ток').AsDouble()
                    cir.LookupParameter('Номер фазы').Set('1, 2, 3')
                elif a == min(0, 0, 0):
                    a += cir.LookupParameter('Ток').AsDouble()
                    cir.LookupParameter('Номер фазы').Set('L1')
                elif b == min(0, 0, 0):
                    b += cir.LookupParameter('Ток').AsDouble()
                    cir.LookupParameter('Номер фазы').Set('L2')
                else:
                    c += cir.LookupParameter('Ток').AsDouble()
                    cir.LookupParameter('Номер фазы').Set('L3')
        phaseShift = (max(a, b, c) - min(a, b, c)) / max(a, b, c) * 100 if any([a, b, c]) else 0
        for cir in pCirs[i]:
                cir.LookupParameter('Перекос фаз').Set(phaseShift)
                cir.LookupParameter('Ток по фазам').Set('{:.2f}, {:.2f}, {:.2f}'.format(float(a), float(b), float(c)).replace('.', ','))

    for panel in pcc.keys():
        appNumber = doc.GetElement(panel).LookupParameter('Номер аппарата защиты').AsInteger()
        for cluster in natural_sorted(pcc[panel].keys(), key=lambda x: x.LookupParameter('Номер группы').AsString()):
            appNumber += 1
            for cir in pcc[panel][cluster]:
                cir.LookupParameter('Номер аппарата защиты').Set(appNumber)

    for panel in pcc.keys():
        for cluster in pcc[panel].keys():
            # print(pcc[panel].keys())
            for cir in pcc[panel][cluster]:
                if 'Щит' in list(cir.Elements)[0].Name:
                    panel2 = list(cir.Elements)[0]
                    fakeCir = filter(lambda x: x.LookupParameter('Тип системы').AsValueString() == 'Данные', panel2.MEPModel.ElectricalSystems)
                    fakeCir = fakeCir[0] if fakeCir else None
                    name = panel2.LookupParameter('Имя щита').AsString()
                    cir.LookupParameter('Имя щита').Set(name if name else '-')
                    cir.LookupParameter('Номер группы').Set('‎') # Там есть невидимый символ
                    cir.LookupParameter('Имя нагрузки').Set('')
                    cir.LookupParameter('Имя нагрузки группы').Set('‎                                    Итого')
                    power1 = 0
                    power2 = 0
                    Q = 0
                    voltage = panel2.LookupParameter('Напряжение щита').AsDouble()
                    U = 220 if voltage == 220 else 380
                    cir.LookupParameter('Напряжение цепи').Set(U)
                    # for cir2 in list(panel2.MEPModel.ElectricalSystems):
                    #     if cir2.Id != cir.Id:
                    #         if fakeCir:
                    #             if cir2.Id != fakeCir.Id:
                    #                 power1 += cir2.LookupParameter('Суммарная мощность группы').AsDouble() / 10763.9104167097
                    #                 power2 += cir2.LookupParameter('Расчетная мощность').AsDouble()
                    #         else:
                    #             power1 += cir2.LookupParameter('Суммарная мощность группы').AsDouble() / 10763.9104167097
                    #             power2 += cir2.LookupParameter('Расчетная мощность').AsDouble()
                    #     Q += cirIds2.LookupParameter('Q, квар').AsDouble()
                    for cir2 in list(panel2.MEPModel.ElectricalSystems):
                        Q += cir2.LookupParameter('Q, квар').AsDouble()
                        if cir2.Id == cir.Id: continue
                        if fakeCir:
                            if cir2.Id == fakeCir.Id: continue
                        if 'Итого' in cir2.LookupParameter('Имя нагрузки группы').AsString():
                            ks = list(cir2.Elements)[0].LookupParameter('Коэффициент спроса расчетный для щита').AsDouble()
                            print(ks)
                            print(cir2.Id)
                            power1 += cir2.LookupParameter('Расчетная мощность').AsDouble()
                            power2 += cir2.LookupParameter('Расчетная мощность').AsDouble() * ks * 1
                        else:
                            power1 += cir2.LookupParameter('Суммарная мощность группы').AsDouble() / 10763.9104167097
                            power2 += cir2.LookupParameter('Расчетная мощность').AsDouble()
                    cir.LookupParameter('Суммарная мощность группы').Set(power1 * 10763.9104167097)
                    cir.LookupParameter('Расчетная мощность').Set(power2)
                    cir.LookupParameter('Q, квар').Set(Q)
                    S = math.sqrt(power2 ** 2 + Q ** 2)
                    cir.LookupParameter('S, кВА').Set(S)
                    I = S / math.sqrt(3) / U * 1000 if U == 380 else S / U * 1000
                    cir.LookupParameter('Iр, А').Set(I)
                    try:
                        ks = power2 / power1
                    except:
                        ks = 0
                    cir.LookupParameter('Коэффициент спроса расчетный').Set(ks)
                    try:
                        cos = power2 / S
                    except:
                        cos = 0
                    cir.LookupParameter('Cos φ').Set(cos)
                    tg = math.tan(math.acos(cos))
                    cir.LookupParameter('tg φ').Set(tg)

                    if fakeCir:
                        name2 = cir.BaseEquipment.LookupParameter('Имя щита').AsString()
                        ks = panel2.LookupParameter('Коэффициент спроса расчетный для щита').AsDouble()
                        fakeCir.LookupParameter('Имя щита').Set(name2 if name2 else '-')
                        fakeCir.LookupParameter('Номер группы').Set(name)
                        fakeCir.LookupParameter('Имя нагрузки').Set(name)
                        fakeCir.LookupParameter('Имя нагрузки группы').Set(name)
                        fakeCir.LookupParameter('Напряжение цепи').Set(U)
                        fakeCir.LookupParameter('Суммарная мощность группы').Set(power2 * 10763.9104167097)
                        fakeCir.LookupParameter('Расчетная мощность').Set(power2 * ks * 1)
                        fakeCir.LookupParameter('Коэффициент спроса расчетный').Set(ks)
                        fakeCir.LookupParameter('Cos φ').Set(cos)
                        fakeCir.LookupParameter('tg φ').Set(tg)

            # spaces = filter(lambda x: '-' not in x, spaces)
            # names = filter(lambda x: 'оробка' not in x, names)
            # for branch in pcb[panel][group]:
            #     for cir in branch:
            #         cir.LookupParameter('Помещения группы').Set('; '.join(natural_sorted(set(spaces))))
            #         cir.LookupParameter('Имя нагрузки группы').Set('; '.join(natural_sorted(set(names))))

###   ###   ###   ###   ###   ###   ###   ###   ###   ###   ###   ###   ###
   ###   ###   ###   ###   ###   ###   ###   ###   ###   ###   ###   ###   ###

t.SetName('ТРН {}, {} с'.format(time.ctime().split()[3], time.time() - startTime))
t.Commit()
