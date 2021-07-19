# -*- coding: utf-8 -*-
""""""
__title__ = 'Расчёт\nспеки'
__author__ = 'SG'


from pyrevit import script
import time
startTime = time.time()
import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *

from Autodesk.Revit.DB import Mechanical, FilteredElementCollector, BuiltInCategory, Transaction, BuiltInParameter, ElementId
from Autodesk.Revit.DB.Plumbing import Pipe

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
k = 304.8
rad = 57.2957795130823
output = script.get_output()

t = Transaction(doc, 'Настройка префиксов и суффиксов воздуховодов')
t.Start()

settings = Mechanical.DuctSettings.GetDuctSettings(doc)
if settings.RoundDuctSizePrefix != 'ø':
    settings.RoundDuctSizePrefix = 'ø'
if settings.RoundDuctSizeSuffix != '':
    settings.RoundDuctSizeSuffix = ''
if settings.RectangularDuctSizeSeparator != '×':
    settings.RectangularDuctSizeSeparator = '×'

t.Commit()

t = Transaction(doc, 'Расчёт спеки')
t.Start()

categories = [
    ('ducts',          'duct_curves',          BuiltInCategory.OST_DuctCurves),          # Воздуховоды
    ('dFits',          'duct_fitting',         BuiltInCategory.OST_DuctFitting),         # Соединительные детали воздуховодов
    ('terms',          'duct_terminal',        BuiltInCategory.OST_DuctTerminal),        # Воздухораспределители
    ('flexes',         'flex_duct_curves',     BuiltInCategory.OST_FlexDuctCurves),      # Гибкие воздуховоды
    ('dArms',          'duct_accessory',       BuiltInCategory.OST_DuctAccessory),       # Арматура воздуховодов
    ('insuls',         'duct_insulations',     BuiltInCategory.OST_DuctInsulations),     # Материалы изоляции воздуховодов
    ('pipes',          'pipe_curves',          BuiltInCategory.OST_PipeCurves),          # Трубы
    ('pFits',          'pipe_fitting',         BuiltInCategory.OST_PipeFitting),         # Соединительные детали трубопроводов
    ('pArms',          'pipe_accessory',       BuiltInCategory.OST_PipeAccessory),       # Арматура трубопроводов
    ('pipeInsuls',     'pipe_insulations',     BuiltInCategory.OST_PipeInsulations),     # Материалы изоляции труб
    ('equipments',     'mechanical_equipment', BuiltInCategory.OST_MechanicalEquipment), # Оборудование
    ('fakes',          'generic_model',        BuiltInCategory.OST_GenericModel),        # Обобщенные модели
    ('piping_systems', 'piping_system',        BuiltInCategory.OST_PipingSystem),        #
    ('sanitary',       'plumbing_fixtures',    BuiltInCategory.OST_PlumbingFixtures),    # Сантехника
    ('_',              'flex_pipe_curves',     BuiltInCategory.OST_FlexPipeCurves),      # Гибким трубы
]

for old_var_name, var_name, bic in categories:
    globals()[var_name] = list(FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType().ToElements())

#2020.04.30 [mechanical_equipment.Add(el) for el in plumbing_fixtures]


def clean_sys_name(sys_name):
    if not sys_name:
        return ''
    return sys_name \
        .split('/')[0] \
        .replace('Х1.', 'Х.') \
        .replace('Х2.', 'Х.') \
        .replace('Х3.', 'Х.') \
        .replace('Х4.', 'Х.') \
        .replace('Хд.', 'Х.')


def set_sys_name(el):
    sys_name = clean_sys_name(el.LookupParameter('Имя системы').AsString())
    el.LookupParameter('ХТ Имя системы').Set(sys_name)
    el.LookupParameter('ХТ Имя системы отступ').Set('‎' + ' ' * 90 + sys_name)


def set_description(el):
    symbol = doc.GetElement(el.GetTypeId())
    description = symbol.LookupParameter('Описание').AsString()
    if not description:
        description = '--- Ошибка. Не заполнено Описание для {}: {}'.format(
            el.LookupParameter('Семейство').AsValueString(),
            el.LookupParameter('Тип').AsValueString(),
        )
    el.LookupParameter('ADSK_Наименование').Set(description)


def get_length_coef(el):
    symbol = doc.GetElement(el.GetTypeId())
    param = symbol.LookupParameter('Код по классификатору').AsString()
    coefficient = float(param) if param else 1.1
    return coefficient


def set_length(el):
    coefficient = get_length_coef(el)
    length = el.LookupParameter('Длина').AsDouble() * k * coefficient
    el.LookupParameter('ХТ Длина ОВ') and el.LookupParameter('ХТ Длина ОВ').Set(length)
    el.LookupParameter('ADSK_Количество') and el.LookupParameter('ADSK_Количество').Set(length)


def print_once(string):
    if 'already_printed' not in globals():
        globals()['bad_symbols'] = {}
    if string not in already_printed:
        print(string)
        already_printed.append(string)


def set_op_rf(el):
    op = doc.GetElement(el.GetTypeId()).LookupParameter('Описание').AsString() or ''
    if el.LookupParameter('ХТ Размер фитинга ОВ'):
        rf = el.LookupParameter('ХТ Размер фитинга ОВ').AsString() or ''
    else:
        rf = ''
    if not el.LookupParameter('оп+рф'):
        print_once('Не найден параметр "оп+рф"')
    el.LookupParameter('оп+рф').Set(op + ' ' + rf)


def set_sort(el):
    if not el.LookupParameter('оп+рф'):
        print_once('Не найден параметр "оп+рф"')
        return
    sys_name = clean_sys_name(el.LookupParameter('Имя системы').AsString())
    if not sys_name:
        continue
    if sys_name[0] == 'П' or sys_name[0] == 'К':
        if sys_name[1].isdigit():
            sort = sys_name[1:] + ' 1 Приток'
        elif sys_name[:2] == 'ПП':
            sort = '600 ПП' + sys_name[2:]
        elif sys_name[:2] == 'ПЕ':
            sort = '700 ПЕ' + sys_name[2:]
        elif sys_name[:2] == 'ПД':
            sort = '800 ПД' + sys_name[2:]
        elif sys_name[:3] == 'Пер':
            sort = '900 ПД' + sys_name[2:]
    elif sys_name[0] == 'В':
        if sys_name[1].isdigit():
            sort = sys_name[1:] + ' 2 Вытяжка'
    elif sys_name[:2] == 'ДУ':
        sort = '500 ДУ' + sys_name[2:]
    elif sys_name[:2] == 'ДП':
        sort = '400 Д' + sys_name[2:] + ' 1 Приток'
    elif sys_name[:2] == 'ДВ':
        sort = '400 Д' + sys_name[2:] + ' 2 Вытяжка'
    elif sys_name[:2] == 'T1':  # Английская
        sort = str(1000 + int(sys_name.split()[-1] if sys_name.split()[-1].isdigit() else '00')) + ' T1'
    elif sys_name[:2] == 'T2':  # Английская
        sort = str(1000 + int(sys_name.split()[-1] if sys_name.split()[-1].isdigit() else '00')) + ' T2'
    elif sys_name[:3] == 'Т1.':  # Русская
        sort = str(1000 + int(sys_name.split()[-1] if sys_name.split()[-1].isdigit() else '00')) + sys_name[2:]
    elif sys_name[:3] == 'Т2.':  # Русская
        sort = str(1000 + int(sys_name.split()[-1] if sys_name.split()[-1].isdigit() else '00')) + sys_name[2:]
    else:
        sort = '9999'
    el.LookupParameter('Сортировка строка').Set(sort)


def check_cost(el):
    cost = doc.GetElement(el.GetTypeId()).LookupParameter('Стоимость').HasValue
    if not cost:
        if 'bad_symbols' not in globals():
            globals()['bad_symbols'] = {}
        definintion = doc.GetElement(el.GetTypeId()).LookupParameter('Описание').AsString() or ''
        if 'Не учитывать' in definintion:
            continue
        bad_symbols[el.GetTypeId()] = (output.linkify(el.Id), el.Id)


def result_check_cost():
    bad_symbols = globals().get('bad_symbols')
    if bad_symbols:
        print('Следует прописать стоимость для сортировки и подсчёта количества:\n{}'.format(
                '\n'.join(['{} {} {} {}'.format(
                    errors[k][0],
                    doc.GetElement(k).LookupParameter('Имя типа').AsString(),
                    doc.GetElement(k).LookupParameter('Имя семейства').AsString(),
                    doc.GetElement(k).LookupParameter('Описание').AsString(),
                    ) for k in errors])
                )
            )
    # el_ids = List[ElementId]([errors[key][1] for key in errors])
    # uidoc.Selection.SetElementIds(el_ids)


# -----------------------------------------------------------------------------------------
# ----------------------------------------------------------------------------------- Трубы
# -----------------------------------------------------------------------------------------

666pipesDict = {
666    6.35 : '1/4" (6,35 mm)',
666    9.52 : '3/8" (9,52 mm)',
666    12.7 : '1/2" (12,7 mm)',
666    15.88: '5/8" (15,9 mm)',
666    15.9 : '5/8" (15,9 mm)',
666    19.05: '3/4" (19,05 mm)',
666    22.23: '7/8" (22,23 mm)',
666    25.4 : '1" (25,4 mm)',
666    28.58: '1-1/8" (28,58 mm)',
666    31.75: '1-1/4" (31,75 mm)',
666    34.93: '1-3/8" (34,93 mm)',
666    41.28: '1-5/8" (41,28 mm)'
666}
666
666for el in pipe_curves:
666    set_sys_name(el)
666
666    if el.LookupParameter('ADSK_Наименование'):
666        symbol = doc.GetElement(el.GetTypeId())
666        description = symbol.LookupParameter('Описание').AsString()
666        if not description or 'егмент' in description:
666            segment_description = el.LookupParameter('Описание сегмента').AsValueString()
666            el.LookupParameter('ADSK_Наименование').Set(segment_description)
666        else:
666            el.LookupParameter('ADSK_Наименование').Set(description)
666
666    size = pipesDict.get(el.LookupParameter('Диаметр').AsDouble() * k)
666    if not size:
666        symbol_name = el.LookupParameter('Тип').AsValueString()
666        segment_name = el.LookupParameter('Сегмент трубы').AsValueString()
666        if '3262-75' in symbol_name:
666            d_inner = el.LookupParameter('Внутренний диаметр').AsDouble() * k
666            d_outer = el.LookupParameter('Внешний диаметр').AsDouble() * k
666            thickness = (d_outer - d_inner) / 2.0
666            size = 'ø{:n}×{:.1f}'.format(size, thickness).replace('.', ',')
666        elif '10704-91' in symbol_name or 'ПВХ Dyka' in segment_name:
666            d_inner = el.LookupParameter('Внутренний диаметр').AsDouble() * k
666            d_outer = el.LookupParameter('Внешний диаметр').AsDouble() * k
666            thickness = (d_outer - d_inner) / 2.0
666            size = '∅{:n}×{:.1f}'.format(d_outer, thickness).replace('.', ',')
666        else:
666            size = 'ø{:.0f}'.format(size)
666    el.LookupParameter('ХТ Размер фитинга ОВ') and el.LookupParameter('ХТ Размер фитинга ОВ').Set(size)
666    el.LookupParameter('ADSK_Марка') and el.LookupParameter('ADSK_Марка').Set(size)
666
666    set_length(el)
666
666    set_op_rf(el)


666# -----------------------------------------------------------------------------------------
666# ----------------------------------------------------------------- Фитинги и арматура труб
666# -----------------------------------------------------------------------------------------
666
666for el in list(pipe_fitting) + list(pipe_accessory):
666    set_sys_name(el)
666
666    set_description(el)
666
666    size = el.LookupParameter('Размер').AsString().replace(',', '.')
666    size_list = sorted([float(el) for el in set(size)], reverse=True) if size[0] else []
666    done_sizes = []
666    for size in size_list:
666        if size in pipesDict:
666            size = pipesDict[size]
666        else:
666            size = 'ø{:.0f}'.format(size)
666            if el.LookupParameter('Угол') and 'ройник' not in el.LookupParameter('Семейство').AsValueString():
666                angle = el.LookupParameter('Угол').AsValueString().split(',')[0]
666                size += ', {}°'.format(angle)
666        done_sizes.append(size)
666    if el.LookupParameter('Артикул') and el.LookupParameter('ADSK_Код изделия'):
666        el.LookupParameter('ADSK_Код изделия').Set(el.LookupParameter('Артикул').AsString())
666    code = doc.GetElement(el.GetTypeId()).LookupParameter('Код по классификатору').AsString()
666    if '0' == code:
666        done_sizes = []
666    komms = doc.GetElement(el.GetTypeId()).LookupParameter('Комментарии к типоразмеру').AsString()
666    komms = komms or ''
666
666    el.LookupParameter('ХТ Размер фитинга ОВ').Set('{} {}'.format(komms, '-'.join(done_sizes)))
666    el.LookupParameter('ADSK_Марка') and el.LookupParameter('ADSK_Марка').Set('{} {}'.format(komms, '-'.join(done_sizes)))
666
666    if 'айка' in el.LookupParameter('Тип').AsValueString() or 'айба' in el.LookupParameter('Тип').AsValueString():
666        el.LookupParameter('ХТ Размер фитинга ОВ').Set('М{:n}'.format(el.LookupParameter('d').AsDouble() * k))  # это зло надо перенести в каркас несущий

666    if el.LookupParameter('ХТ Длина ОВ'):
666        param = el.LookupParameter('Длина')
666        coefficient = get_length_coef(el)
666        length = round(param.AsDouble() * k * coefficient, 2) if param else 0
666        el.LookupParameter('ХТ Длина ОВ').Set(length)
666        el.LookupParameter('ADSK_Количество') and el.LookupParameter('ADSK_Количество').Set(length)

666    if 'чтено в' in code:
666        el.LookupParameter('ХТ Размер фитинга ОВ').Set('')
666        el.LookupParameter('ADSK_Примечание').Set(code) if el.LookupParameter('ADSK_Примечание') else None
666    else:
666        code and code != '0' and el.LookupParameter('ADSK_Код изделия').Set(code)
666
666
666    set_op_rf(el)

666# -----------------------------------------------------------------------------------------
666# ---------------------------------------------------------------------------- Оборудование
666# -----------------------------------------------------------------------------------------
666
666for el in mechanical_equipment + plumbing_fixtures:
666    if el.LookupParameter('Категория').AsValueString() == 'Сантехнические приборы':
666        el.LookupParameter('ХТ Имя системы').Set('В1')
666    else:
666        set_sys_name(el)
666
666    set_description(el)
666
666    typeMarka = doc.GetElement(el.GetTypeId()).LookupParameter('Комментарии к типоразмеру').AsString()
666    el.LookupParameter('ХТ Размер фитинга ОВ') and el.LookupParameter('ХТ Размер фитинга ОВ').Set(typeMarka)
666    el.LookupParameter('ADSK_Марка') and el.LookupParameter('ADSK_Марка').Set(typeMarka)
666
666    param = doc.GetElement(el.GetTypeId()).LookupParameter('Код по классификатору').AsString()
666    param and el.LookupParameter('ADSK_Код изделия').Set(param)
666
666    if el.GetSubComponentIds():
666        parent_sys_name = clean_sys_name(el.LookupParameter('ХТ Имя системы').AsString())
666        for sub_el_id in el.GetSubComponentIds():
666            doc.GetElement(sub_el_id).LookupParameter('ХТ Имя системы').Set(parent_sys_name)
666
666    set_op_rf(el)


# -----------------------------------------------------------------------------------------
# --------------------------------------------------------------------------- Изоляция труб
# -----------------------------------------------------------------------------------------

666assortment = {
666    15: '13×22 (для Ду15)',
666    20: '13×28 (для Ду20)',
666    25: '13×35 (для Ду25)',
666    32: '13×42 (для Ду32)',
666    40: '13×48 (для Ду40)',
666    50: '13×60 (для Ду50)',
666    65: '13×76 (для Ду65)',
666    80: '13×89 (для Ду80)',
666    100: '13×108 (для Ду100)',
666    125: '13×133 (для Ду125)',
666    150: '13×160 (для Ду150)',
666    0: 'Рулон 13x1000',
666}
666
666for el in pipe_insulations:
666    set_sys_name(el)
666
666    set_description(el)
666
666    host = doc.GetElement(el.HostElementId)
666    if isinstance(host, Pipe):
666        coefficient = get_length_coef(el)
666        length = el.LookupParameter('Длина').AsDouble() * k * coefficient
666
666        diameter = int(host.LookupParameter('Диаметр').AsDouble() * k)
666        size = assortment.get(diameter)
666        if not size:
666            size = 'Не найден: {}'.format(diameter)
666            print('Для {} не найден диаметр изоляции: {}'.format(output.linkify(el.Id), diameter))
666    else:
666        diameter = int(host.LookupParameter('Размер').AsString().split('-')[0])
666        length = diameter * 2.5 / k
666        size = assortment.get(diameter, 'Не найден: {}'.format(diameter))
666
666    komms = doc.GetElement(el.GetTypeId()).LookupParameter('Комментарии к типоразмеру').AsString()
666    el.LookupParameter('ХТ Размер фитинга ОВ').Set((komms + ' ' if komms else '') + size)
666    el.LookupParameter('ADSK_Марка') and el.LookupParameter('ADSK_Марка').Set((komms + ' ' if komms else '') + size)
666    el.LookupParameter('ХТ Длина ОВ').Set(length)
666    el.LookupParameter('ADSK_Количество') and el.LookupParameter('ADSK_Количество').Set(length)
666
666    set_op_rf(el)
666
666
666# -----------------------------------------------------------------------------------------
666# ---------------------------------------------------------------------------- Гибкие трубы
666# -----------------------------------------------------------------------------------------
666
666for el in flex_pipe_curves:
666    set_sys_name(el)
666
666    set_description(el)
666
666    diameter = el.LookupParameter('Диаметр').AsDouble() * k
666    el.LookupParameter('ADSK_Марка') and el.LookupParameter('ADSK_Марка').Set('ø{}'.format(diameter))
666
666    set_length(el)
666
666    set_op_rf(el)
666
666
666# ------------------------------------------------------------------------------------------------------------------------------
666#  Воздуховоды -----------------------------------------------------------------------------------------------------------------
666# ------------------------------------------------------------------------------------------------------------------------------
666
666for el in duct_curves:
666    set_sys_name(el)
666
666    set_description(el)
666
666    set_length(el)
666
666    if duct.LookupParameter('Ширина'):                                         # Прямоугольные
666        b = round(el.LookupParameter('Ширина').AsDouble() * k, 2)
666        h = round(el.LookupParameter('Высота').AsDouble() * k, 2)
666        if h > b:
666            b, h = h, b
666        if 'Д' in el.LookupParameter('ХТ Имя системы').AsString():             # Возможны проблемы с определением противодымки
666            mark = 'δ=1,2'
666        else:
666            mark = 'δ=0,5' if b < 251 else 'δ=0,7' if b < 1001 else 'δ=0,9'
666        mark = '{:.0f}×{:.0f}, {}'.format(b, h, mark)
666    else:                                                                      # Круглые
666        d = round(el.LookupParameter('Диаметр').AsDouble() * k, 2)
666        if 'Д' in el.LookupParameter('ХТ Имя системы').AsString():
666            mark = 'δ=1,2'
666        else:
666            mark = 'δ=0,5' if d < 201 else 'δ=0,6' if d < 451 else 'δ=0,7'
666        mark = 'ø{:.0f}, {}'.format(d, mark)
666    if length < 10:
666        mark = 'Не учитывать ' + mark
666    el.LookupParameter('ADSK_Марка') and el.LookupParameter('ADSK_Марка').Set('ø{}'.format(mark))
666
666    set_op_rf(el)
666
666
666# ------------------------------------------------------------------------------------------------------------------------------
666#  Гибкие воздуховоды ----------------------------------------------------------------------------------------------------------
666# ------------------------------------------------------------------------------------------------------------------------------
666
666symbols = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FlexDuctCurves).WhereElementIsElementType().ToElements()
666symbols = [el for el in symbols if el.LookupParameter('Имя типа').AsString() == 'SG_Круглый изолированный']
666symbol = symbols[0] if symbols else None
666error = False
666for el in flex_duct_curves:
666    set_sys_name(el)
666
666    set_description(el)
666
666    mark = 'ø{:.0f}'.format(el.LookupParameter('Диаметр').AsDouble() * k)
666    el.LookupParameter('ХТ Размер фитинга ОВ') and el.LookupParameter('ХТ Размер фитинга ОВ').Set(mark)
666    el.LookupParameter('ADSK_Марка') and el.LookupParameter('ADSK_Марка').Set(mark)
666
666    if el.LookupParameter('Классификация систем').AsString() == 'Приточный воздух' \
666            and el.LookupParameter('Тип').AsValueString() != 'SG_Круглый изолированный':
666        if symbol:
666            el.FlexDuctType = symbol
666        else:
666            error = True
666
666    set_length(el)
666
666    set_op_rf(el)
666if error:
666    print('Не найден тип гибкого воздуховода "SG_Круглый изолированный"')
666
666
666# ------------------------------------------------------------------------------------------------------------------------------
666#  Материалы изоляции воздуховодов ---------------------------------------------------------------------------------------------
666# ------------------------------------------------------------------------------------------------------------------------------
666
666kk = 10.76391041671
666kkk = k * kk / 1000
666for el in duct_insulations:
666    set_sys_name(el)
666
666    set_description(el)
666
666    host = doc.GetElement(el.HostElementId)
666    coefficient = get_length_coef(el)
666    if host.LookupParameter('ХТ Площадь'):
666        area = host.LookupParameter('ХТ Площадь').AsDouble() / kkk * coefficient
666    elif el.LookupParameter('Площадь') and el.LookupParameter('Площадь').AsDouble() > 0:
666        area = el.LookupParameter('Площадь').AsDouble() / kkk * coefficient
666    else:
666        raise Exception('Надо обработать этот вариант')
666    el.LookupParameter('ХТ Длина ОВ') and el.LookupParameter('ХТ Длина ОВ').Set(area)
666    el.LookupParameter('ADSK_Количество') and el.LookupParameter('ADSK_Количество').Set(area)
666
666    insType = el.Name
666    if host.LookupParameter('Комментарии').AsString() == 'Утепление+':
666        mark = 'δ=32'
666    elif 'пожар' in insType or 'гнеза' in insType:
666        koms = doc.GetElement(el.GetTypeId()).LookupParameter('Комментарии к типоразмеру').AsString()
666        if koms:
666            mark = koms
666    else:
666        mark = 'δ=10'
666    el.LookupParameter('ХТ Размер фитинга ОВ') and el.LookupParameter('ХТ Размер фитинга ОВ').Set(mark)
666    el.LookupParameter('ADSK_Марка') and el.LookupParameter('ADSK_Марка').Set(mark)
666
666    thickness = 10 / k if mark == 'δ=10' else 32 / k
666    pos.LookupParameter('Толщина изоляции').Set(thickness)
666
666    set_op_rf(el)
666
666
666# ------------------------------------------------------------------------------------------------------------------------------
666#  Соединительные детали воздуховодов + Арматура воздуховодов ------------------------------------------------------------------
666# ------------------------------------------------------------------------------------------------------------------------------
666
666for el in duct_fitting + duct_accessory:
666    set_sys_name(el)
666
666    set_description(el)
666
666    size = el.LookupParameter('Размер').AsString()
666    size.replace(' мм', '')
666    if size[:(size.find("-"))] == size[(size.find("-") + 1):]:
666        size = size[:(size.find("-"))]
666
666    if el.LookupParameter('Смещение S'):
666        offset = el.LookupParameter('Смещение S').AsDouble()
666        offset = offset * k
666        size = '{}, S={:.0f}'.format(size, offset)
666
666    if el.LookupParameter('Угол отвода'):
666        angle = el.LookupParameter('Угол отвода').AsDouble() * rad
666        if angle > 0.5:
666            size = '{}, {:.0f}°'.format(size, angle)
666#   typeMarka = doc.GetElement(el.GetTypeId()).LookupParameter('Тип, марка, обозначение документа, опросного листа').AsString() Тут место для Маркировка типоразмера
666#   if typeMarka: size = typeMarka
666
666#  -------------------------------------------------------------- Обработать круглые врезки
666
666    name = doc.GetElement(el.GetTypeId()).get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
666    if 'ФБО' in name:
666        size = name
666    elif 'НП' in name:
666        size = name
666    opis = doc.GetElement(el.GetTypeId()).LookupParameter('Описание').AsString()
666    if opis:
666        if 'Переход стальной с прямоугольного на прямоугольное сечение' in opis:
666            x1 = round(el.LookupParameter('Ширина воздуховода 1').AsDouble() * k, 2)
666            x2 = round(el.LookupParameter('Ширина воздуховода 2').AsDouble() * k, 2)
666            y1 = round(el.LookupParameter('Высота воздуховода 1').AsDouble() * k, 2)
666            y2 = round(el.LookupParameter('Высота воздуховода 2').AsDouble() * k, 2)
666            dx = round(el.LookupParameter('ШиринаСмещения').AsDouble() * k, 2)
666            dy = round(el.LookupParameter('ВысотаСмещения').AsDouble() * k, 2)
666
666            X1 = x1
666            X2 = x2
666            Y1 = y1
666            Y2 = y2
666            DX = dx
666            DY = dy
666            f1 = X1 * Y1
666            f2 = X2 * Y2
666            width = X1 >= X2
666            maxf = f1 >= f2
666            height = Y1 >= Y2
666            total = width + maxf + height
666            revX = 1
666            if total == 1:
666                revX = -1
666            if total == 2:
666                revX = -1
666            revY = -1
666            if f1 >= f2:
666                revY = 1
666
666            if abs(DX - X1 / 2) < 2:
666                xxx = 'center'  # Допустимая погрешность 2 мм
666            elif abs(DX - X1 / 2 + abs(X1 - X2) / 2 * revX) < 2:
666                xxx = 'left'
666            elif abs(DX - X1 / 2 - abs(X1 - X2) / 2 * revX) < 2:
666                xxx = 'right'
666            else:
666                xxx = 'error'
666
666            if abs(DY - Y1 / 2) < 2:
666                yyy = 'center'
666            elif abs(DY - Y1 / 2 + abs(Y1 - Y2) / 2 * revY) < 2:
666                yyy = 'up'
666            elif abs(DY - Y1 / 2 - abs(Y1 - Y2) / 2 * revY) < 2:
666                yyy = 'down'
666            else:
666                yyy = 'error'
666
666            x, y = xxx, yyy
666
666            if y == 'up':
666                if x == 'left': type_ = 'Тип 1'
666                if x == 'center': type_ = 'Тип 2'
666                if x == 'right': type_ = 'Тип 3'
666                if x == 'error': type_ = 'Error'
666            if y == 'down':
666                if x == 'left': type_ = 'Тип 3'
666                if x == 'center': type_ = 'Тип 2'
666                if x == 'right': type_ = 'Тип 1'
666                if x == 'error': type_ = 'Error'
666            if y == 'center':
666                if x == 'left': type_ = 'Тип 2 поворот'
666                if x == 'center': type_ = 'Тип 4'
666                if x == 'right': type_ = 'Тип 2 поворот'
666                if x == 'error': type_ = 'Error'
666            if y == 'error': type_ = 'Error'
666
666            sizeAsList = size.replace('-', '×').split('×')
666
666            try:
666                x1 = sizeAsList[0]
666                y1 = sizeAsList[1]
666                x2 = sizeAsList[2]
666                y2 = sizeAsList[3]
666            except:
666                type_ = 'Error:Equal'
666                x2 = '0'
666                y2 = '0'
666
666            if 'поворот' in type_:
666                size = y1 + '×' + x1 + '-' + y2 + '×' + x2 + ', Тип 2'
666            else:
666                size = x1 + '×' + y1 + '-' + x2 + '×' + y2 + ', ' + type_
666
666    el.LookupParameter('ХТ Размер фитинга ОВ') and el.LookupParameter('ХТ Размер фитинга ОВ').Set(size)
666    el.LookupParameter('ADSK_Марка') and el.LookupParameter('ADSK_Марка').Set(size)
666
666    set_op_rf(el)


# ------------------------------------------------------------------------------------------------------------------------------
#  Воздухораспределители -------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------------------------------------------------

donelist = []
for el in duct_terminal:
    set_sys_name(el)

    set_description(el)

    if el.LookupParameter('Семейство').AsValueString() == 'Шкаф вытяжной с ребрами':
        h = el.LookupParameter('Смещение').AsDouble()
        # try:
        #     el.LookupParameter('Высота потолка').Set(h if h * k > 2000 else 2400 / k)
        # except AttributeError:
        #     pass
        el.LookupParameter('Высота потолка') and el.LookupParameter('Высота потолка').Set(h if h * k > 2000 else 2400 / k)
        symbol = doc.GetElement(el.GetTypeId())
        if symbol.Id not in donelist:
            height_bottom = int(symbol.LookupParameter('Высота нижнего отверстия').AsDouble() * k)
            height_bottom = height_bottom if height_bottom > 100 else 100
            height_top = int(symbol.LookupParameter('Высота верхнего отверстия').AsDouble() * k)
            height_top = height_top if height_top > 100 else 100
            name = str(int(symbol.LookupParameter('Ширина шкафа').AsDouble() * k))
            name += '×' + str(int(symbol.LookupParameter('Глубина шкафа').AsDouble() * k))
            name += '; Отверстия: ▲' + str(int(symbol.LookupParameter('Ширина верхнего отверстия').AsDouble() * k))
            name += '×' + str(height_top)
            name += 'h, ▼' + str(int(symbol.LookupParameter('Ширина нижнего отверстия').AsDouble() * k))
            name += '×' + str(height_bottom)
            name += 'h;'
            name += '\nполная высота стены {:.0f} мм'.format(el.LookupParameter('Смещение').AsDouble() * k)
            # symbol.LookupParameter('Группа модели').Set(name)
            donelist.append(symbol.Id)

            el.LookupParameter('ХТ Размер фитинга ОВ') and el.LookupParameter('ХТ Размер фитинга ОВ').Set(name)
            el.LookupParameter('ADSK_Марка') and el.LookupParameter('ADSK_Марка').Set(name)

    set_op_rf(el)

t.SetName('Расчёт спеки {}, Δ={} с'.format(time.ctime().split()[3], time.time() - startTime))
t.Commit()