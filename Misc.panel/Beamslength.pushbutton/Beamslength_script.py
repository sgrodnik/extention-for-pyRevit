# -*- coding: utf-8 -*-
""""""
__title__ = 'КД'
__author__ = 'SG'
import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

k = 304.8
k1 = 35.31466672149
k2 = 1000000 / k**2  # 10.763910416709722

els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).WhereElementIsNotElementType().ToElements()

els = [i for i in els if i.LookupParameter('Использование в конструкции').AsValueString() != 'Горизонтальный раскос']

t = Transaction(doc, 'КД')

t.Start()

# Заполнение наименований экземпляров Каркаса несущего (КН)
for el in els:
    dimL = el.LookupParameter('Фактическая длина').AsDouble() * k
    el.LookupParameter('Длина факт').Set('{:.0f}'.format(dimL))
    dimB = el.LookupParameter('ADSK_Размер_Высота').AsDouble() * k
    dimH = el.LookupParameter('ADSK_Размер_Ширина').AsDouble() * k
    name = '{:.0f}×{:.0f}'.format(min([dimB, dimH]), max([dimB, dimH]))
    el.LookupParameter('Наименование').Set(name)
    el.LookupParameter('Наименование СМ').Set('Пиломатериал ' + name)
    el.LookupParameter('Наименование краткое').Set(name)
    zapas = 1.0 if 'Кле' in el.LookupParameter('Тип').AsValueString() else 1.1
    el.LookupParameter('ХТ Длина ОВ').Set(dimL * dimB * dimH / 1000000000 * zapas)
    # else:
    #     dimL = el.LookupParameter('Фактическая длина').AsDouble() * k
    #     el.LookupParameter('ХТ Длина ОВ').Set(dimL / 1000)


    if not el.LookupParameter('Этап').HasValue:
        el.LookupParameter('Этап').Set(999 * k2)

# Сортировка
els = [el for el in sorted(els, key=lambda x: x.LookupParameter('Фактическая длина').AsDouble(), reverse=True)]
els = [el for el in sorted(els, key=lambda x: float('.'.join(map(lambda x: '{:0>5}'.format(x), x.LookupParameter('Наименование').AsString().split('×')))), reverse=True)]
els = [el for el in sorted(els, key=lambda x: x.LookupParameter('Этап').AsDouble())]

# Простановка Позиции
i = 0
part = 0
name = 0
length = 0
for el in els:
    if 'Не СЭ' in el.LookupParameter('Тип').AsValueString():
        continue
    newpart = el.LookupParameter('Этап').AsDouble()
    newname = el.LookupParameter('Наименование').AsString()
    newlength = round(el.LookupParameter('Фактическая длина').AsDouble() * k, 0)
    if part != newpart or name != newname or length != newlength:
        i += 1
        part, name, length = newpart, newname, newlength
    el.LookupParameter('ADSK_Позиция').Set('{}'.format(i))

# # Простановка Позиции без дублей
# i = 0
# part = 0
# name = '0'
# length = 0
# individuals = {}
# for el in els:
#     if 'Не СЭ' in el.LookupParameter('Тип').AsValueString():
#         continue
#     newpart = el.LookupParameter('Этап').AsDouble()
#     newname = el.LookupParameter('Наименование').AsString()
#     newlength = round(el.LookupParameter('Фактическая длина').AsDouble() * k, 0)
#     individ = newname + ' ' + str(newlength)
#     if individ not in individuals.keys():
#         if part != newpart or name != newname or length != newlength:
#             i += 1
#             individuals[individ] = i
#             part, name, length = newpart, newname, newlength
#         el.LookupParameter('ADSK_Позиция').Set('{}'.format(i))
#     else:
#         el.LookupParameter('ADSK_Позиция').Set('{}'.format(individuals[individ]))


# Расчёт объёма каждого экземпляра и формирования словаря словарей объёмов
values = {}
for el in els:
    part = el.LookupParameter('Этап').AsDouble()
    if part not in values.keys():
        values[part] = {}
    name = el.LookupParameter('Наименование').AsString()
    if name not in values[part].keys():
        values[part][name] = []
    a = el.LookupParameter('ADSK_Размер_Высота').AsDouble()
    b = el.LookupParameter('ADSK_Размер_Ширина').AsDouble()
    c = el.LookupParameter('Фактическая длина').AsDouble()
    value = a * b * c
    values[part][name].append(value)

q_params = {
    'q1': 1 * k2,
    'q2': 2 * k2,
    'q3': 3 * k2,
    'q4': 4 * k2,
    'q5': 5 * k2,
    'q6': 6 * k2,
    'q7': 7 * k2,
    'q8': 8 * k2,
    'q9': 9 * k2,
}

# Подсчёт КН по этапам и суммарный, а также запись
for el in els:
    name = el.LookupParameter('Наименование').AsString()

    for q_param in q_params.keys():
        if q_params[q_param] in values.keys():
            if name in values[q_params[q_param]].keys():
                el.LookupParameter(q_param).Set('{:.3f}'.format(sum(values[q_params[q_param]][name]) / k1).replace('.', ','))
              # el.LookupParameter('q1')   .Set('{:.3f}'.format(sum(values[1 * k2           ][name]) / k1).replace('.', ','))  <-<-<- Example -<-<-<
            else:
                el.LookupParameter(q_param).Set('')
              # el.LookupParameter('q1'   ).Set('')  <-<-<- Example -<-<-<

    total = 0
    for q_param in sorted(q_params.keys()):
        if q_params[q_param] in values.keys():
            if name in values[q_params[q_param]].keys():
                total += sum(values[q_params[q_param]][name])
                # print(total / k1)
    el.LookupParameter('q999').Set('{:.3f}'.format(total / k1).replace('.', ','))

partNames = {
    1 * k2: '1 Обвязка',
    2 * k2: '2 Балки',
    3 * k2: '3 Лаги',
    4 * k2: '4 Контрутепление',
    5 * k2: '5 Мауэрлат',
    6 * k2: '6 Подстропильные балки',
    7 * k2: '7 Накосные',
    8 * k2: '8 Стропила',
    9 * k2: '9 Лобовая доска',
}

# Конвертация числового этапа в строковый
for el in els:
    partCode = el.LookupParameter('Этап').AsDouble()
    if partCode in partNames.keys():
        el.LookupParameter('Этап строка').Set(partNames[partCode])


#######################################################################################################################################################
########################################################## Соединения несущего каркаса ################################################################
#######################################################################################################################################################


els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructConnections).WhereElementIsNotElementType().ToElements()

# Заполнение наименований экземпляров Соединений несущего каркаса (СНК)
for el in els:
    param = doc.GetElement(el.GetTypeId()).LookupParameter('Описание')
    description = param.AsString() if param.HasValue else ''
    param = doc.GetElement(el.GetTypeId()).LookupParameter('Комментарии к типоразмеру')
    comms = param.AsString() if param.HasValue else ''
    param = doc.GetElement(el.GetTypeId()).LookupParameter('Группа модели')
    group = param.AsString() if param.HasValue else ''
    if description == 'Шпилька резьбовая':
        diam = el.LookupParameter('D').AsDouble() * k
        length = el.LookupParameter('L').AsDouble() * k + el.LookupParameter('b').AsDouble() * k
        comms = 'М{:.0f}'.format(diam)
        el.LookupParameter('Длина факт').Set('{:.0f}'.format(length))
        el.LookupParameter('Наименование').Set('{} {}, L{:.0f}'.format(description, comms, length))
        el.LookupParameter('Наименование СМ').Set('{} {}'.format(description, comms))
    else:
        el.LookupParameter('Наименование').Set('{} {}'.format(description, comms))
        el.LookupParameter('Наименование СМ').Set('{} {}'.format(description, comms))
    el.LookupParameter('Наименование краткое').Set('{} {}'.format(group, comms))
    # el.LookupParameter('Наименование краткое').Set('{}'.format(group))

    if not el.LookupParameter('Этап').HasValue:
        el.LookupParameter('Этап').Set(999 * k2)

# Сортировка
els = [el for el in sorted(els, key=lambda x: x.LookupParameter('Наименование').AsString())]
els = [el for el in sorted(els, key=lambda x: x.LookupParameter('Этап').AsDouble())]

# Формирование словаря словарей, содержащих экземпляры СНК
structCons = {}
for el in els:
    part = el.LookupParameter('Этап').AsDouble()
    if part not in structCons.keys():
        structCons[part] = {}
    name = el.LookupParameter('Наименование СМ').AsString()
    if name not in structCons[part].keys():
        structCons[part][name] = []
    structCons[part][name].append(el)

# print(structCons.keys())

# Подсчёт СНК по этапам и суммарный, а также запись
for el in els:
    name = el.LookupParameter('Наименование СМ').AsString()

    for q_param in sorted(q_params.keys()):
        # print(q_param)
        if q_params[q_param] in structCons.keys():
            # print(1)
            if name in structCons[q_params[q_param]].keys():
                # print('{}'.format(len(structCons[q_params[q_param]][name])))
                if structCons[q_params[q_param]][name][0].Name == 'Закладная шпилька':
                    el.LookupParameter(q_param).Set('{:.1f}'.format(sum(map(lambda x: x.LookupParameter('L').AsDouble() * k + x.LookupParameter('b').AsDouble() * k, structCons[q_params[q_param]][name])) / 1000).replace('.', ','))
                else:
                    el.LookupParameter(q_param).Set('{}'.format(len(structCons[q_params[q_param]][name])))
              # el.LookupParameter('q1'   ).Set('{}'.format(len(structCons[1 * k2]           [name])).replace('.', ','))  <-<-<- Example -<-<-<
            else:
                # print(el.Id)
                el.LookupParameter(q_param).Set('')
              # el.LookupParameter('q1'   ).Set('')  <-<-<- Example -<-<-<

    total = 0
    for q_param in q_params.keys():
        if q_params[q_param] in structCons.keys():
            # print(123)
            if name in structCons[q_params[q_param]].keys():
                if structCons[q_params[q_param]][name][0].Name == 'Закладная шпилька':
                    total += sum(map(lambda x: x.LookupParameter('L').AsDouble() * k + x.LookupParameter('b').AsDouble() * k, structCons[q_params[q_param]][name])) / 1000
                else:
                    total += len(structCons[q_params[q_param]][name])
    # total = round()
    if isinstance(total, int):
        el.LookupParameter('q999').Set('{}'.format(total))
    else:
        el.LookupParameter('q999').Set('{:.1f}'.format(total).replace('.', ','))

# Конвертация числового этапа в строковый
for el in els:
    partCode = el.LookupParameter('Этап').AsDouble()
    if partCode in partNames.keys():
        el.LookupParameter('Этап строка').Set(partNames[partCode])

# Подсчёт Крыш
areaOSB = 0
volumeUtepl = 0
areaSves = 0
areaOtdelki = 0
areaGidro = 0
areaGidroSten = 0
roofs = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Roofs).WhereElementIsNotElementType().ToElements()
for roof in roofs:
    if roof.Name == 'Плита OSB-3, 18 мм':
        areaOSB += roof.LookupParameter('Площадь').AsDouble()
    elif roof.Name == 'Утеплитель':
        volumeUtepl += roof.LookupParameter('Объем').AsDouble()
    elif roof.Name == 'Площадь подшивки свесов':
        areaSves += roof.LookupParameter('Площадь').AsDouble()
    elif roof.Name == 'Площадь отделки':
        areaOtdelki += roof.LookupParameter('Площадь').AsDouble()
    elif roof.Name == 'Гидроизоляция обвязки':
        areaGidro += roof.LookupParameter('Площадь').AsDouble()
    elif roof.Name == 'Гидроизоляция стен':
        areaGidroSten += roof.LookupParameter('Площадь').AsDouble()

    if not roof.LookupParameter('Этап').HasValue:
        roof.LookupParameter('Этап').Set(999 * k2)


# Подсчёт Лобовых досок
lengthKon = 0.0
lengthNak = 0.0
lengthSves = 0.0
fascias = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Fascia).WhereElementIsNotElementType().ToElements()
for fascia in fascias:
    if fascia.Name == 'Длина конька':
        lengthKon += fascia.LookupParameter('Длина').AsDouble()
    elif fascia.Name == 'Длина накосных':
        lengthNak += fascia.LookupParameter('Длина').AsDouble()
    elif fascia.Name == 'Длина свесов':
        lengthSves += fascia.LookupParameter('Длина').AsDouble()

# Запись подсчитанного в фейки
fakes = [el for el in els if el.LookupParameter('Семейство').AsValueString() == 'Фейк']

for fake in fakes:
    if fake.LookupParameter('Тип').AsValueString() == 'Длина конька':
        fake.LookupParameter('q999').Set('{:.1f}'.format(lengthKon * k / 1000).replace('.', ','))
        fake.LookupParameter('ХТ Длина ОВ').Set(lengthKon * k / 1000 * 1.1)

    elif fake.LookupParameter('Тип').AsValueString() == 'Длина накосных':
        fake.LookupParameter('q999').Set('{:.1f}'.format(lengthNak * k / 1000).replace('.', ','))
        fake.LookupParameter('ХТ Длина ОВ').Set(lengthNak * k / 1000 * 1.1)

    elif fake.LookupParameter('Тип').AsValueString() == 'Длина свесов':
        fake.LookupParameter('q999').Set('{:.1f}'.format(lengthSves * k / 1000).replace('.', ','))
        fake.LookupParameter('ХТ Длина ОВ').Set(lengthSves * k / 1000 * 1.1)

    elif fake.LookupParameter('Тип').AsValueString() == 'Плита OSB-3, 18 мм':
        fake.LookupParameter('q9').Set('{:.0f}'.format(areaOSB / k2).replace('.', ','))
        fake.LookupParameter('q999').Set('{:.0f}'.format(areaOSB / k2).replace('.', ','))
        fake.LookupParameter('ХТ Длина ОВ').Set(areaOSB / k2 * 1.1)

    elif fake.LookupParameter('Тип').AsValueString() == 'Утеплитель 50':
        fake.LookupParameter('q999').Set('{:.0f}'.format(volumeUtepl / k1).replace('.', ','))
        fake.LookupParameter('ХТ Длина ОВ').Set(volumeUtepl / k1 * 1.1 / 3.0 * 1.0)

    elif fake.LookupParameter('Тип').AsValueString() == 'Утеплитель 100':
        fake.LookupParameter('q999').Set('{:.0f}'.format(volumeUtepl / k1).replace('.', ','))
        fake.LookupParameter('ХТ Длина ОВ').Set(volumeUtepl / k1 * 1.1 / 3.0 * 2.0)

    elif fake.LookupParameter('Тип').AsValueString() == 'Площадь кровли':
        fake.LookupParameter('q999').Set('{:.0f}'.format(areaOSB / k2 * 1.1).replace('.', ','))
        fake.LookupParameter('ХТ Длина ОВ').Set(areaOSB / k2 * 1.1)

    elif fake.LookupParameter('Тип').AsValueString() == 'Подкладочный ковёр под битумную черепицу':
        fake.LookupParameter('q999').Set('{:.0f}'.format(areaOSB / k2 * 1.1).replace('.', ','))
        fake.LookupParameter('ХТ Длина ОВ').Set(areaOSB / k2 * 1.1)

    elif fake.LookupParameter('Тип').AsValueString() == 'Площадь подшивки свесов':
        fake.LookupParameter('q999').Set('{:.0f}'.format(areaSves / k2 * 1.1).replace('.', ','))
        fake.LookupParameter('ХТ Длина ОВ').Set(areaSves / k2 * 1.1)

    elif fake.LookupParameter('Тип').AsValueString() == 'Площадь отделки':
        fake.LookupParameter('q999').Set('{:.0f}'.format(areaOtdelki / k2 * 1.1).replace('.', ','))
        fake.LookupParameter('ХТ Длина ОВ').Set(areaOtdelki / k2 * 1.1)

    elif fake.LookupParameter('Тип').AsValueString() == 'Гидроизоляция обвязки':
        fake.LookupParameter('q999').Set('{:.0f}'.format(areaGidro / k2 * 1.1).replace('.', ','))
        fake.LookupParameter('ХТ Длина ОВ').Set(areaGidro / k2 * 1.1)

    elif fake.LookupParameter('Тип').AsValueString() == 'Гидроизоляция стен':
        fake.LookupParameter('q999').Set('{:.0f}'.format(areaGidroSten / k2 * 1.1).replace('.', ','))
        fake.LookupParameter('ХТ Длина ОВ').Set(areaGidroSten / k2 * 1.1)

t.Commit()
