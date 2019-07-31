# -*- coding: utf-8 -*-
""""""
__title__ = 'КД'
__author__ = 'SG'
import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

k = 304.8
k1 = 35.31466672149
k2 = 1000000 / k**2  # 10.763910416709722

els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).WhereElementIsNotElementType().ToElements()

t = Transaction(doc, 'КД')
t.Start()

for el in els:
    el.LookupParameter('Длина факт').Set('{:.0f}'.format(el.LookupParameter('Фактическая длина').AsDouble() * k))
    dimB = el.LookupParameter('ADSK_Размер_Высота').AsDouble() * k
    dimH = el.LookupParameter('ADSK_Размер_Ширина').AsDouble() * k
    el.LookupParameter('Наименование').Set('{:.0f}×{:.0f}'.format(min([dimB, dimH]), max([dimB, dimH])))

    if not el.LookupParameter('Этап').HasValue:
        el.LookupParameter('Этап').Set(999 * k2)

# i = 1
# dct = {}
# for el in els:
#     part = el.LookupParameter('Этап').AsDouble()
#     if part not in dct.keys():
#         dct[part] = {}
#     dct[part] = {}
#     name = el.LookupParameter('Наименование').AsString()
#     if name not in dct[part].keys():
#         dct[part][name] = []
#     dct[part][name].append(el)
#     i += 1


els = [el for el in sorted(els, key=lambda x: x.LookupParameter('Фактическая длина').AsDouble(), reverse=True)]
els = [el for el in sorted(els, key=lambda x: float('.'.join(map(lambda x: '{:0>5}'.format(x), x.LookupParameter('Наименование').AsString().split('×')))), reverse=True)]
els = [el for el in sorted(els, key=lambda x: x.LookupParameter('Этап').AsDouble())]
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
    'q4': 4 * k2,
    'q5': 5 * k2,
    'q6': 6 * k2,
    'q7': 7 * k2,
    'q8': 8 * k2,
    'q9': 9 * k2,
}

for el in els:
    part = el.LookupParameter('Этап').AsDouble()
    name = el.LookupParameter('Наименование').AsString()

    for q_param in q_params.keys():
        if name in values[q_params[q_param]].keys():
            el.LookupParameter(q_param).Set('{:.3f}'.format(sum(values[q_params[q_param]][name]) / k1).replace('.', ','))
          # el.LookupParameter('q1')   .Set('{:.3f}'.format(sum(values[1 * k2           ][name]) / k1).replace('.', ','))  <-<-<-Example-<-<-<
        else:
            el.LookupParameter(q_param).Set('')
          # el.LookupParameter('q1'   ).Set('')  <-<-<-Example-<-<-<

    total = sum(
        [sum(values[1 * k2][name] if name in values[1 * k2].keys() else [0]),
         sum(values[2 * k2][name] if name in values[2 * k2].keys() else [0]),
         sum(values[3 * k2][name] if name in values[3 * k2].keys() else [0]),
         sum(values[4 * k2][name] if name in values[4 * k2].keys() else [0]),
         sum(values[5 * k2][name] if name in values[5 * k2].keys() else [0]),
         sum(values[6 * k2][name] if name in values[6 * k2].keys() else [0]),
         sum(values[7 * k2][name] if name in values[7 * k2].keys() else [0]),
         sum(values[8 * k2][name] if name in values[8 * k2].keys() else [0]),
         sum(values[9 * k2][name] if name in values[9 * k2].keys() else [0])]
    ) / k1
    el.LookupParameter('q999').Set('{:.3f}'.format(total).replace('.', ','))

# for el in els:
#     part = el.LookupParameter('Этап').AsDouble()
#     name = el.LookupParameter('Наименование').AsString()
#     try:
#         el.LookupParameter('q1').Set('{:.3f}'.format(sum(values[1 * k2][name]) / k1).replace('.', ','))
#     except KeyError:
#         el.LookupParameter('q1').Set('')
#     try:
#         el.LookupParameter('q2').Set('{:.3f}'.format(sum(values[2 * k2][name]) / k1).replace('.', ','))
#     except KeyError:
#         el.LookupParameter('q2').Set('')
#     try:
#         el.LookupParameter('q3').Set('{:.3f}'.format(sum(values[3 * k2][name]) / k1).replace('.', ','))
#     except KeyError:
#         el.LookupParameter('q3').Set('')
#     try:
#         el.LookupParameter('q4').Set('{:.3f}'.format(sum(values[4 * k2][name]) / k1).replace('.', ','))
#     except KeyError:
#         el.LookupParameter('q4').Set('')
#     try:
#         el.LookupParameter('q5').Set('{:.3f}'.format(sum(values[5 * k2][name]) / k1).replace('.', ','))
#     except KeyError:
#         el.LookupParameter('q5').Set('')
#     try:
#         el.LookupParameter('q6').Set('{:.3f}'.format(sum(values[6 * k2][name]) / k1).replace('.', ','))
#     except KeyError:
#         el.LookupParameter('q6').Set('')
#     try:
#         el.LookupParameter('q7').Set('{:.3f}'.format(sum(values[7 * k2][name]) / k1).replace('.', ','))
#     except KeyError:
#         el.LookupParameter('q7').Set('')
#     try:
#         el.LookupParameter('q8').Set('{:.3f}'.format(sum(values[8 * k2][name]) / k1).replace('.', ','))
#     except KeyError:
#         el.LookupParameter('q8').Set('')
#     try:
#         el.LookupParameter('q9').Set('{:.3f}'.format(sum(values[9 * k2][name]) / k1).replace('.', ','))
#     except KeyError:
#         el.LookupParameter('q9').Set('')
#     try:
#         total = sum(
#             [sum(values[1 * k2][name] if name in values[1 * k2].keys() else [0]),
#              sum(values[2 * k2][name] if name in values[2 * k2].keys() else [0]),
#              sum(values[3 * k2][name] if name in values[3 * k2].keys() else [0]),
#              sum(values[4 * k2][name] if name in values[4 * k2].keys() else [0]),
#              sum(values[5 * k2][name] if name in values[5 * k2].keys() else [0]),
#              sum(values[6 * k2][name] if name in values[6 * k2].keys() else [0]),
#              sum(values[7 * k2][name] if name in values[7 * k2].keys() else [0]),
#              sum(values[8 * k2][name] if name in values[8 * k2].keys() else [0]),
#              sum(values[9 * k2][name] if name in values[9 * k2].keys() else [0])]
#         ) / k1
#         el.LookupParameter('q999').Set('{:.3f}'.format(total).replace('.', ','))
#     except KeyError:
#         el.LookupParameter('q999').Set('')

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

for el in els:
    partCode = el.LookupParameter('Этап').AsDouble()
    if partCode in partNames.keys():
        el.LookupParameter('Этап строка').Set(partNames[partCode])

els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructConnections).WhereElementIsNotElementType().ToElements()

for el in els:
    param = doc.GetElement(el.GetTypeId()).LookupParameter('Описание')
    description = param.AsString() if param.HasValue else ''
    param = doc.GetElement(el.GetTypeId()).LookupParameter('Комментарии к типоразмеру')
    comms = param.AsString() if param.HasValue else ''
    el.LookupParameter('Наименование').Set('{} {}'.format(description, comms))

    if not el.LookupParameter('Этап').HasValue:
        el.LookupParameter('Этап').Set(999 * k2)

els = [el for el in sorted(els, key=lambda x: x.LookupParameter('Наименование').AsString())]
els = [el for el in sorted(els, key=lambda x: x.LookupParameter('Этап').AsDouble())]

structCons = {}
for el in els:
    part = el.LookupParameter('Этап').AsDouble()
    if part not in structCons.keys():
        structCons[part] = {}
    name = el.LookupParameter('Наименование').AsString()
    if name not in structCons[part].keys():
        structCons[part][name] = []
    structCons[part][name].append(el)

for el in els:
    part = el.LookupParameter('Этап').AsDouble()
    name = el.LookupParameter('Наименование').AsString()
    try:
        el.LookupParameter('q1').Set('{}'.format(len(structCons[1 * k2][name])).replace('.', ','))
    except KeyError:
        el.LookupParameter('q1').Set('')
    try:
        el.LookupParameter('q2').Set('{}'.format(len(structCons[2 * k2][name])).replace('.', ','))
    except KeyError:
        el.LookupParameter('q2').Set('')
    try:
        el.LookupParameter('q3').Set('{}'.format(len(structCons[3 * k2][name])).replace('.', ','))
    except KeyError:
        el.LookupParameter('q3').Set('')
    try:
        el.LookupParameter('q4').Set('{}'.format(len(structCons[4 * k2][name])).replace('.', ','))
    except KeyError:
        el.LookupParameter('q4').Set('')
    try:
        el.LookupParameter('q5').Set('{}'.format(len(structCons[5 * k2][name])).replace('.', ','))
    except KeyError:
        el.LookupParameter('q5').Set('')
    try:
        el.LookupParameter('q6').Set('{}'.format(len(structCons[6 * k2][name])).replace('.', ','))
    except KeyError:
        el.LookupParameter('q6').Set('')
    try:
        el.LookupParameter('q7').Set('{}'.format(len(structCons[7 * k2][name])).replace('.', ','))
    except KeyError:
        el.LookupParameter('q7').Set('')
    try:
        el.LookupParameter('q8').Set('{}'.format(len(structCons[8 * k2][name])).replace('.', ','))
    except KeyError:
        el.LookupParameter('q8').Set('')
    try:
        el.LookupParameter('q9').Set('{}'.format(len(structCons[9 * k2][name])).replace('.', ','))
    except KeyError:
        el.LookupParameter('q9').Set('')

    total = sum(
        [len(structCons[1 * k2][name] if 1 * k2 in structCons.keys() and name in structCons[1 * k2].keys() else []),
         len(structCons[2 * k2][name] if 2 * k2 in structCons.keys() and name in structCons[2 * k2].keys() else []),
         len(structCons[3 * k2][name] if 3 * k2 in structCons.keys() and name in structCons[3 * k2].keys() else []),
         len(structCons[4 * k2][name] if 4 * k2 in structCons.keys() and name in structCons[4 * k2].keys() else []),
         len(structCons[5 * k2][name] if 5 * k2 in structCons.keys() and name in structCons[5 * k2].keys() else []),
         len(structCons[6 * k2][name] if 6 * k2 in structCons.keys() and name in structCons[6 * k2].keys() else []),
         len(structCons[7 * k2][name] if 7 * k2 in structCons.keys() and name in structCons[7 * k2].keys() else []),
         len(structCons[8 * k2][name] if 8 * k2 in structCons.keys() and name in structCons[8 * k2].keys() else []),
         len(structCons[9 * k2][name] if 9 * k2 in structCons.keys() and name in structCons[9 * k2].keys() else []),
         ])
    el.LookupParameter('q999').Set('{}'.format(total).replace('.', ','))


for el in els:
    partCode = el.LookupParameter('Этап').AsDouble()
    if partCode in partNames.keys():
        el.LookupParameter('Этап строка').Set(partNames[partCode])

areaOSB = 0
volumeUtepl = 0
roofs = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Roofs).WhereElementIsNotElementType().ToElements()
for roof in roofs:
    if roof.Name == 'ОСП18':
        areaOSB += roof.LookupParameter('Площадь').AsDouble()
    elif roof.Name == 'Утеплитель':
        volumeUtepl += roof.LookupParameter('Объем').AsDouble()


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

fakes = [el for el in els if el.LookupParameter('Семейство').AsValueString() == 'Фейк']

for fake in fakes:
    if fake.LookupParameter('Тип').AsValueString() == 'Длина конька':
        fake.LookupParameter('q999').Set('{:.1f}'.format(lengthKon * k / 1000).replace('.', ','))
    elif fake.LookupParameter('Тип').AsValueString() == 'Длина накосных':
        fake.LookupParameter('q999').Set('{:.1f}'.format(lengthNak * k / 1000).replace('.', ','))
    elif fake.LookupParameter('Тип').AsValueString() == 'Длина свесов':
        fake.LookupParameter('q999').Set('{:.1f}'.format(lengthSves * k / 1000).replace('.', ','))
    elif fake.LookupParameter('Тип').AsValueString() == 'ОСП18':
        fake.LookupParameter('q9').Set('{:.0f}'.format(areaOSB / k2).replace('.', ','))
        fake.LookupParameter('q999').Set('{:.0f}'.format(areaOSB / k2).replace('.', ','))
    elif fake.LookupParameter('Тип').AsValueString() == 'Утеплитель':
        fake.LookupParameter('q999').Set('{:.0f}'.format(volumeUtepl / k1).replace('.', ','))
    elif fake.LookupParameter('Тип').AsValueString() == 'Площадь кровли':
        fake.LookupParameter('q999').Set('{:.0f}'.format(areaOSB / k2 * 1.1).replace('.', ','))
    elif fake.LookupParameter('Тип').AsValueString() == 'Подкладочный ковёр под битумную черепицу':
        fake.LookupParameter('q999').Set('{:.0f}'.format(areaOSB / k2 * 1.1).replace('.', ','))

t.Commit()
