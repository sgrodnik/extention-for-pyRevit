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
from System import Guid

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
k = 304.8
kk = 10.76391041671
kkk = k * kk / 1000
rad = 57.2957795130823
TO_DEGREE = 1 / 1.5707963267949 * 90
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

categories = [
    ('ducts',          'duct_curves',          BuiltInCategory.OST_DuctCurves,          1),  # Воздуховоды
    ('dFits',          'duct_fitting',         BuiltInCategory.OST_DuctFitting,         1),  # Соединительные детали воздуховодов
    ('terms',          'duct_terminal',        BuiltInCategory.OST_DuctTerminal,        1),  # Воздухораспределители
    ('flexes',         'flex_duct_curves',     BuiltInCategory.OST_FlexDuctCurves,      1),  # Гибкие воздуховоды
    ('dArms',          'duct_accessory',       BuiltInCategory.OST_DuctAccessory,       1),  # Арматура воздуховодов
    ('insuls',         'duct_insulations',     BuiltInCategory.OST_DuctInsulations,     1),  # Материалы изоляции воздуховодов
    ('pipes',          'pipe_curves',          BuiltInCategory.OST_PipeCurves,          1),  # Трубы
    ('pFits',          'pipe_fitting',         BuiltInCategory.OST_PipeFitting,         1),  # Соединительные детали трубопроводов
    ('pArms',          'pipe_accessory',       BuiltInCategory.OST_PipeAccessory,       1),  # Арматура трубопроводов
    ('pipeInsuls',     'pipe_insulations',     BuiltInCategory.OST_PipeInsulations,     1),  # Материалы изоляции труб
    ('equipments',     'mechanical_equipment', BuiltInCategory.OST_MechanicalEquipment, 1),  # Оборудование
    ('fakes',          'generic_model',        BuiltInCategory.OST_GenericModel,        1),  # Обобщенные модели
    # ('piping_systems', 'piping_system',        BuiltInCategory.OST_PipingSystem,      1),  #
    ('sanitary',       'plumbing_fixtures',    BuiltInCategory.OST_PlumbingFixtures,    1),  # Сантехника
    ('_',              'flex_pipe_curves',     BuiltInCategory.OST_FlexPipeCurves,      1),  # Гибким трубы
]

elems = []
for old_var_name, var_name, bic, calc in categories:
    globals()[var_name] = list(FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType().ToElements())
    if not calc:
        globals()[var_name] = []
    else:
        elems += globals()[var_name]


def clean_sys_name(sys_name):
    if not sys_name:
        return ''
    if 'Т1.1' in sys_name or 'Т2.1' in sys_name: return 'Т1.1-Т2.1'
    if 'Т1.2' in sys_name or 'Т2.2' in sys_name: return 'Т1.2-Т2.2'
    if 'Т1.3' in sys_name or 'Т2.3' in sys_name: return 'Т1.3-Т2.3'
    if 'Т1.4' in sys_name or 'Т2.4' in sys_name: return 'Т1.4-Т2.4'
    if 'Т1.5' in sys_name or 'Т2.5' in sys_name: return 'Т1.5-Т2.5'
    if 'Т1.6' in sys_name or 'Т2.6' in sys_name: return 'Т1.6-Т2.6'
    if 'Т1.7' in sys_name or 'Т2.7' in sys_name: return 'Т1.7-Т2.7'
    if 'Т1.8' in sys_name or 'Т2.8' in sys_name: return 'Т1.8-Т2.8'
    return sys_name \
        .split('/')[0] \
        .replace('Х1.', 'Х.') \
        .replace('Х2.', 'Х.') \
        .replace('Х3.', 'Х.') \
        .replace('Х4.', 'Х.') \
        .replace('Хд.', 'Х.') \
        .split(',')[0]


def set_sys_name(el):
    if not el.LookupParameter('Имя системы'):
        return
    elif el in plumbing_fixtures:
        sys_name = 'В1'
    else:
        if hasattr(el, 'SuperComponent') and el.SuperComponent:
            sys_name = clean_sys_name(el.SuperComponent.LookupParameter('Имя системы').AsString())
        elif not el.LookupParameter('Имя системы').AsString():
            sys_name = 'Не определено'
        else:
            sys_name = clean_sys_name(el.LookupParameter('Имя системы').AsString())
    el.LookupParameter('ХТ Имя системы').Set(sys_name)
    el.LookupParameter('ХТ Имя системы отступ') and el.LookupParameter('ХТ Имя системы отступ').Set('‎' + ' ' * 90 + sys_name)


def set_description(el):
    symbol = doc.GetElement(el.GetTypeId())
    description = symbol.LookupParameter('Описание').AsString() or ''
    if not description:
        description = '--- Ошибка. Не заполнено Описание для {}: {}'.format(
            el.LookupParameter('Семейство').AsValueString(),
            el.LookupParameter('Тип').AsValueString(),
        )
    elif el in pipe_curves and 'егмент' in description:
        segment_description = el.LookupParameter('Описание сегмента').AsString()
        try:
            el.LookupParameter('ADSK_Наименование').Set(segment_description)
        except:
            param = el.get_Parameter(Guid('e6e0f5cd-3e26-485b-9342-23882b20eb43'))
            param.Set(segment_description)
        return
    # el.LookupParameter('ADSK_Наименование').Set(description)
    param = el.get_Parameter(Guid('e6e0f5cd-3e26-485b-9342-23882b20eb43'))
    if param and not param.IsReadOnly:
        param.Set(description)


def get_mark_pipe_curves(el):
    size = el.LookupParameter('Диаметр').AsDouble() * k
    size = pipe_assortment.get(size)
    if not size:
        size = el.LookupParameter('Диаметр').AsDouble() * k
        symbol_name = el.LookupParameter('Тип').AsValueString()
        segment_name = el.LookupParameter('Сегмент трубы').AsValueString()
        if ' ду' in symbol_name:
            size = 'ø{:.0f}'.format(size)
        elif '3262-75' in symbol_name:
            d_inner = el.LookupParameter('Внутренний диаметр').AsDouble() * k
            d_outer = el.LookupParameter('Внешний диаметр').AsDouble() * k
            thickness = (d_outer - d_inner) / 2.0
            size = 'ø{:n}×{:.1f}'.format(size, thickness).replace('.', ',')
        elif '10704-91' in symbol_name or 'ПВХ Dyka' in segment_name or '10704-91' in segment_name:
            d_inner = el.LookupParameter('Внутренний диаметр').AsDouble() * k
            d_outer = el.LookupParameter('Внешний диаметр').AsDouble() * k
            thickness = (d_outer - d_inner) / 2.0
            size = '∅{:n}×{:.1f}'.format(d_outer, thickness).replace('.', ',')
        elif 'KAN_PE-Xc' in symbol_name:
            d_inner = el.LookupParameter('Внутренний диаметр').AsDouble() * k
            d_outer = el.LookupParameter('Внешний диаметр').AsDouble() * k
            thickness = (d_outer - d_inner) / 2.0
            size = '{:n}×{:.1f}'.format(d_outer, thickness).replace('.', ',')
        else:
            size = 'ø{:.0f}'.format(size)
    return size


pipe_assortment = {
    6.35 : '1/4" (6,35 mm)',
    9.52 : '3/8" (9,52 mm)',
    12.7 : '1/2" (12,7 mm)',
    15.88: '5/8" (15,9 mm)',
    15.9 : '5/8" (15,9 mm)',
    19.05: '3/4" (19,05 mm)',
    22.23: '7/8" (22,23 mm)',
    25.4 : '1" (25,4 mm)',
    28.58: '1-1/8" (28,58 mm)',
    31.75: '1-1/4" (31,75 mm)',
    34.93: '1-3/8" (34,93 mm)',
    41.28: '1-5/8" (41,28 mm)'
}
pipe_assortment2 = {  # KAN_Push - Отвод гибкий PE-Xc
    12: '12×2,0',
    14: '14×2,0',
    18: '18×2,5',
    25: '25×3,5',
}


def get_mark_pipe_fitting_pipe_accessory(el):
    symbol_name = el.LookupParameter('Тип').AsValueString()
    family_name = el.LookupParameter('Семейство').AsValueString()
    if 'айка' in symbol_name or 'айба' in symbol_name:
        mark = 'М{:n}'.format(el.LookupParameter('d').AsDouble() * k)  # это зло надо перенести в каркас несущий
    elif 'PE-Xc' in symbol_name and 'KAN_Push - Отвод гибкий' in family_name:
        size = el.LookupParameter('Размер').AsString().replace(',', '.').replace('ø', '').split('-')[0]
        mark = pipe_assortment2[float(size)]
    else:
        size_list = el.LookupParameter('Размер').AsString().replace(',', '.').replace('ø', '').split('-')
        size_list = sorted([float(i) for i in set(size_list)], reverse=True) if size_list[0] else []
        done_sizes = []
        for size in size_list:
            if size in pipe_assortment:
                size = pipe_assortment[size]
            else:
                size = 'ø{:.0f}'.format(size)
                if el.LookupParameter('Угол') and 'ройник' not in el.LookupParameter('Семейство').AsValueString():
                    angle = el.LookupParameter('Угол').AsValueString().split(',')[0]
                    size += ', {}°'.format(angle)
            done_sizes.append(size)
        code = doc.GetElement(el.GetTypeId()).LookupParameter('Код по классификатору').AsString() or ''
        if '0' == code:
            done_sizes = []
        komms = doc.GetElement(el.GetTypeId()).LookupParameter('Комментарии к типоразмеру').AsString() or ''
        mark = '{} {}'.format(komms, '-'.join(done_sizes))
    return mark


pipe_insulation_assortment = {  # K-Flex ST
    15: '13×22 (для Ду15)',
    20: '13×28 (для Ду20)',
    25: '13×35 (для Ду25)',
    32: '13×42 (для Ду32)',
    40: '13×48 (для Ду40)',
    50: '13×60 (для Ду50)',
    65: '13×76 (для Ду65)',
    80: '13×89 (для Ду80)',
    100: '13×108 (для Ду100)',
    125: '13×133 (для Ду125)',
    150: '13×160 (для Ду150)',
    0: 'Рулон 13x1000',
}
pipe_insulation_assortment = {  # Тизол EURO-ШЕЛЛ Ц 80
    15: '30×22 (для Ду15)',
    20: '30×28 (для Ду20)',
    25: '30×35 (для Ду25)',
    32: '30×42 (для Ду32)',
    40: '30×48 (для Ду40)',
    50: '30×60 (для Ду50)',
    65: '30×76 (для Ду65)',
    80: '30×89 (для Ду80)',
    100: '30×108 (для Ду100)',
    125: '30×133 (для Ду125)',
    125.0: '30×133 (для Ду125)',
    150: '30×160 (для Ду150)',
    200: '30×219 (для Ду200)',
    250: '30×273 (для Ду250)',
}
pipe_insulation_assortment2 = {  # K-Flex PE COMPACT
    12: '6×15 (для 12×2,0)',
    14: '6×15 (для 14×2,0)',
    18: '6×18 (для 18×2,5)',
    25: '6×28 (для 25×3,5)',
}
pipe_insulation_assortment3 = {
    6.35 : '⌀=6, δ=9',
    9.52 : '⌀=10, δ=9',
    12.7 : '⌀=12, δ=9',
    15.88 : '⌀=15, δ=13',
    15.9 : '⌀=15, δ=13',
    19.05: '⌀=18, δ=13',
    22.23: '⌀=22, δ=13',
    25.4 : '⌀=25, δ=13',
    28.58: '⌀=28, δ=13',
    34.93: '⌀=35, δ=13',
    41.28: '⌀=42, δ=13',
    15: '⌀=22, δ=9',
    20: '⌀=28, δ=9',
    25: '⌀=35, δ=9',
    32: '⌀=42, δ=9',
    40: '⌀=48, δ=9',
    50: '⌀=60, δ=9',
    65: '⌀=76, δ=9',
    80: '⌀=89, δ=9',
    90: '⌀=102, δ=9',
    100:'⌀=114, δ=9',
    125:'⌀=140, δ=13',
    150:'⌀=160, δ=13'
}


def get_mark_pipe_insulations(el):
    host = doc.GetElement(el.HostElementId)
    if 'в стяжке' in el.Name:
        assortment = pipe_insulation_assortment2
    else:
        assortment = pipe_insulation_assortment
    if isinstance(host, Pipe):
        diameter = int(host.LookupParameter('Диаметр').AsDouble() * k)
        size = assortment.get(diameter)
        if not size:
            size = pipe_insulation_assortment3.get(diameter)
            if not size:
                size = 'Ошибка: Не найден ø{}'.format(diameter)
                print('Для {} не найден диаметр изоляции: {}'.format(output.linkify(el.Id), diameter))
    else:
        diameter = int(host.LookupParameter('Размер').AsString().split('-')[0])
        length = diameter * 2.5 / k
        size = assortment.get(diameter, 'Ошибка: Не найден ø{}'.format(diameter))

    komms = doc.GetElement(el.GetTypeId()).LookupParameter('Комментарии к типоразмеру').AsString()
    mark = (komms + ' ' if komms else '') + size
    return mark


def get_mark_duct_curves(el):
    if el.LookupParameter('Ширина'):                                         # Прямоугольные
        b = round(el.LookupParameter('Ширина').AsDouble() * k, 2)
        h = round(el.LookupParameter('Высота').AsDouble() * k, 2)
        if h > b:
            b, h = h, b
        if 'Д' in el.LookupParameter('ХТ Имя системы').AsString():             # Возможны проблемы с определением противодымки
            mark = 'δ=1,2'
        else:
            mark = 'δ=0,5' if b < 251 else 'δ=0,7' if b < 1001 else 'δ=0,9'
        mark = '{:.0f}×{:.0f}, {}'.format(b, h, mark)
    else:                                                                      # Круглые
        d = round(el.LookupParameter('Диаметр').AsDouble() * k, 2)
        if 'Д' in el.LookupParameter('ХТ Имя системы').AsString():
            mark = 'δ=1,2'
        else:
            mark = 'δ=0,5' if d < 201 else 'δ=0,6' if d < 451 else 'δ=0,7'
        mark = 'ø{:.0f}, {}'.format(d, mark)
    length = el.LookupParameter('Длина').AsDouble() * k
    if length < 10:
        mark = 'Не учитывать ' + mark
    mark = 'ø{}'.format(mark)
    return mark


def get_mark_duct_insulations(el):
    host = doc.GetElement(el.HostElementId)
    insType = el.Name
    if host.LookupParameter('Комментарии').AsString() == 'Утепление+':
        mark = 'δ=32'
    elif 'пожар' in insType or 'гнеза' in insType:
        mark = doc.GetElement(el.GetTypeId()).LookupParameter('Комментарии к типоразмеру').AsString() or ''
    else:
        mark = 'δ=10'
    return mark


def get_mark_duct_fitting_duct_accessory(el):
    size = el.LookupParameter('Размер').AsString()
    size.replace(' мм', '')
    if size[:(size.find("-"))] == size[(size.find("-") + 1):]:
        size = size[:(size.find("-"))]

    if el.LookupParameter('Смещение S'):
        offset = el.LookupParameter('Смещение S').AsDouble()
        offset = offset * k
        size = '{}, S={:.0f}'.format(size, offset)

    if el.LookupParameter('Угол отвода'):
        angle = el.LookupParameter('Угол отвода').AsDouble() * rad
        if angle > 0.5:
            size = '{}, {:.0f}°'.format(size, angle)

    name = doc.GetElement(el.GetTypeId()).get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    if 'ФБО' in name:
        size = name
    elif 'НП' in name:
        size = name
    opis = doc.GetElement(el.GetTypeId()).LookupParameter('Описание').AsString() or ''
    if 'Переход стальной с прямоугольного на прямоугольное сечение' in opis:
        x1 = round(el.LookupParameter('Ширина воздуховода 1').AsDouble() * k, 2)
        x2 = round(el.LookupParameter('Ширина воздуховода 2').AsDouble() * k, 2)
        y1 = round(el.LookupParameter('Высота воздуховода 1').AsDouble() * k, 2)
        y2 = round(el.LookupParameter('Высота воздуховода 2').AsDouble() * k, 2)
        dx = round(el.LookupParameter('ШиринаСмещения').AsDouble() * k, 2)
        dy = round(el.LookupParameter('ВысотаСмещения').AsDouble() * k, 2)

        X1 = x1
        X2 = x2
        Y1 = y1
        Y2 = y2
        DX = dx
        DY = dy
        f1 = X1 * Y1
        f2 = X2 * Y2
        width = X1 >= X2
        maxf = f1 >= f2
        height = Y1 >= Y2
        total = width + maxf + height
        revX = 1
        if total == 1:
            revX = -1
        if total == 2:
            revX = -1
        revY = -1
        if f1 >= f2:
            revY = 1

        if abs(DX - X1 / 2) < 2:
            xxx = 'center'  # Допустимая погрешность 2 мм
        elif abs(DX - X1 / 2 + abs(X1 - X2) / 2 * revX) < 2:
            xxx = 'left'
        elif abs(DX - X1 / 2 - abs(X1 - X2) / 2 * revX) < 2:
            xxx = 'right'
        else:
            xxx = 'error'

        if abs(DY - Y1 / 2) < 2:
            yyy = 'center'
        elif abs(DY - Y1 / 2 + abs(Y1 - Y2) / 2 * revY) < 2:
            yyy = 'up'
        elif abs(DY - Y1 / 2 - abs(Y1 - Y2) / 2 * revY) < 2:
            yyy = 'down'
        else:
            yyy = 'error'

        x, y = xxx, yyy

        if y == 'up':
            if x == 'left': type_ = 'Тип 1'
            if x == 'center': type_ = 'Тип 2'
            if x == 'right': type_ = 'Тип 3'
            if x == 'error': type_ = 'Error'
        if y == 'down':
            if x == 'left': type_ = 'Тип 3'
            if x == 'center': type_ = 'Тип 2'
            if x == 'right': type_ = 'Тип 1'
            if x == 'error': type_ = 'Error'
        if y == 'center':
            if x == 'left': type_ = 'Тип 2 поворот'
            if x == 'center': type_ = 'Тип 4'
            if x == 'right': type_ = 'Тип 2 поворот'
            if x == 'error': type_ = 'Error'
        if y == 'error': type_ = 'Error'

        sizeAsList = size.replace('-', '×').split('×')

        try:
            x1 = sizeAsList[0]
            y1 = sizeAsList[1]
            x2 = sizeAsList[2]
            y2 = sizeAsList[3]
        except:
            type_ = 'Error:Equal'
            x2 = '0'
            y2 = '0'

        if 'поворот' in type_:
            size = y1 + '×' + x1 + '-' + y2 + '×' + x2 + ', Тип 2'
        else:
            size = x1 + '×' + y1 + '-' + x2 + '×' + y2 + ', ' + type_
    return size


def get_mark_duct_terminal(el):
    if el.LookupParameter('Семейство').AsValueString() == 'Шкаф вытяжной с ребрами':
        h = el.LookupParameter('Смещение').AsDouble()
        el.LookupParameter('Высота потолка') and el.LookupParameter('Высота потолка').Set(h if h * k > 2000 else 2400 / k)
        symbol = doc.GetElement(el.GetTypeId())
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
    else:
        name = doc.GetElement(el.GetTypeId()).LookupParameter('Группа модели').AsString() or ''
    return name


def set_mark(el):
    if el in pipe_curves:
        mark = get_mark_pipe_curves(el)
    elif el in pipe_fitting + pipe_accessory:
        mark = get_mark_pipe_fitting_pipe_accessory(el)
    elif el in mechanical_equipment + plumbing_fixtures:
        mark = doc.GetElement(el.GetTypeId()).LookupParameter('Комментарии к типоразмеру').AsString() or ''
    elif el in pipe_insulations:
        mark = get_mark_pipe_insulations(el)
    elif el in flex_pipe_curves:
        #2020.05.09 mark = 'ø{}'.format(el.LookupParameter('Диаметр').AsDouble() * k)
        mark = 'ø{:n}'.format(el.LookupParameter('Диаметр').AsDouble() * k)
    elif el in duct_curves:
        mark = get_mark_duct_curves(el)
    elif el in flex_duct_curves:
        #2020.05.09 mark = 'ø{:.0f}'.format(el.LookupParameter('Диаметр').AsDouble() * k)
        mark = 'ø{:n}'.format(el.LookupParameter('Диаметр').AsDouble() * k)  # Объединить с flex_pipe_curves?
    elif el in duct_insulations:
        mark = get_mark_duct_insulations(el)
        thickness = 10 / k if mark == 'δ=10' else 32 / k
        el.LookupParameter('Толщина изоляции').Set(thickness)
    elif el in duct_fitting + duct_accessory:
        mark = get_mark_duct_fitting_duct_accessory(el)
    elif el in duct_terminal:
        mark = get_mark_duct_terminal(el)
    else:
        mark = ''

    el.LookupParameter('ХТ Размер фитинга ОВ') and (not el.LookupParameter('ХТ Размер фитинга ОВ').IsReadOnly) and el.LookupParameter('ХТ Размер фитинга ОВ').Set(mark)
    el.LookupParameter('ADSK_Марка') and (not el.LookupParameter('ADSK_Марка').IsReadOnly) and el.LookupParameter('ADSK_Марка').Set(mark)


def get_coef(el):
    default = 1.1
    symbol = doc.GetElement(el.GetTypeId())
    code = symbol.LookupParameter('Код по классификатору').AsString().replace(',', '.') or ''
    if '((' not in code:
        return default
    code = code.split('((')[1].split('))')[0]
    coefficient = float(code) if code else default
    return coefficient


def set_amount(el):
    coefficient = get_coef(el)


    if el in pipe_insulations:
        host = doc.GetElement(el.HostElementId)
        if isinstance(host, Pipe):
            length = el.LookupParameter('Длина').AsDouble() * k * coefficient
        else:
            diameter = int(host.LookupParameter('Размер').AsString().split('-')[0])
            length = diameter * 2.5 * coefficient

    elif el in duct_insulations:
        host = doc.GetElement(el.HostElementId)
        if host.LookupParameter('ХТ Площадь'):
            area = host.LookupParameter('ХТ Площадь').AsDouble() / kkk * coefficient
        elif el.LookupParameter('Площадь') and el.LookupParameter('Площадь').AsDouble() > 0:
            area = el.LookupParameter('Площадь').AsDouble() / 13.1870819297193 * 1.225 * 1000 * coefficient
        else:
            # raise Exception('Надо обработать этот вариант')
            area = 0
        length = area

    else:
        symbol_name = el.LookupParameter('Тип').AsValueString()
        family_name = el.LookupParameter('Семейство').AsValueString()
        if 'PE-Xc' in symbol_name and 'KAN_Push - Отвод гибкий' in family_name:
            radius = el.LookupParameter('H').AsDouble() * k
            angle = el.LookupParameter('Угол').AsDouble() * TO_DEGREE
            length = 3.14159 * radius * angle / 180
        elif 'Фейк' in family_name:
            return
        else:
            if el.LookupParameter('Комментарии').AsString() and 'Особая длина' in el.LookupParameter('Комментарии').AsString():
                length = el.LookupParameter('Комментарии').AsString().split()[-1]
                length = float(length)
            else:
                param = el.LookupParameter('Длина')
                length = param.AsDouble() * k * coefficient if param else 0

    el.LookupParameter('ХТ Длина ОВ') and el.LookupParameter('ХТ Длина ОВ').Set(length / k)
    el.LookupParameter('ADSK_Количество') and el.LookupParameter('ADSK_Количество').Set(length)


def set_code(el):
    code = doc.GetElement(el.GetTypeId()).LookupParameter('Код по классификатору').AsString() or ''
    code = code.split('((')[0] if '((' in code else code
    if 'чтено в' in code:
        el.LookupParameter('ХТ Размер фитинга ОВ').Set('')
        el.LookupParameter('ADSK_Примечание') and el.LookupParameter('ADSK_Примечание').Set(code)
    else:
        code != '0' and el.LookupParameter('ADSK_Код изделия') and not el.LookupParameter('ADSK_Код изделия').IsReadOnly and el.LookupParameter('ADSK_Код изделия').Set(code)
    if el.LookupParameter('Артикул') and el.LookupParameter('ADSK_Код изделия') and not el.LookupParameter('ADSK_Код изделия').IsReadOnly:
        el.LookupParameter('ADSK_Код изделия').Set(el.LookupParameter('Артикул').AsString())


def set_subcomponents(el):
    if hasattr(el, 'GetSubComponentIds'):
        if el.GetSubComponentIds():
            parent_sys_name = clean_sys_name(el.LookupParameter('ХТ Имя системы').AsString())
            for sub_el_id in el.GetSubComponentIds():
                param = doc.GetElement(sub_el_id).LookupParameter('ХТ Имя системы')
                param and param.Set(parent_sys_name)


def print_once(string):
    if 'already_printed' not in globals():
        globals()['already_printed'] = []
    if string not in already_printed:
        print(string)
        already_printed.append(string)


def set_op_rf(el):
    if el.LookupParameter('ADSK_Наименование'):
        naim = el.LookupParameter('ADSK_Наименование').AsString() or ''
    else:
        param = el.get_Parameter(Guid('e6e0f5cd-3e26-485b-9342-23882b20eb43'))
        if param:
            naim = param.AsString() or ''
        else:
            naim = doc.GetElement(el.GetTypeId()).LookupParameter('Описание').AsString() or ''
    if el.LookupParameter('ADSK_Марка'):
        mark = el.LookupParameter('ADSK_Марка').AsString() or ''
    elif el.LookupParameter('ХТ Размер фитинга ОВ'):
        mark = el.LookupParameter('ХТ Размер фитинга ОВ').AsString() or ''
    else:
        mark = ''
    if el.LookupParameter('ADSK_Код изделия'):
        code = el.LookupParameter('ADSK_Код изделия').AsString() or ''
    else:
        code = ''
    el.LookupParameter('наим марка код') and el.LookupParameter('наим марка код').Set(naim + ' | ' + mark + ' | ' + code)
    op = doc.GetElement(el.GetTypeId()).LookupParameter('Описание').AsString() or ''
    if el.LookupParameter('ХТ Размер фитинга ОВ'):
        rf = el.LookupParameter('ХТ Размер фитинга ОВ').AsString() or ''
    else:
        rf = ''
    if not el.LookupParameter('оп+рф'):
        if not el.LookupParameter('наим марка код'):
            print_once('Не найден параметр "оп+рф" ("наим марка код")')
        # else:
        #     el.LookupParameter('наим марка код').Set(op + ' ' + rf)  # Переделать в н-м-к!!!!!!!!!!!!!!!!!!!!!!!!!!!!2020.05.11
    else:
        el.LookupParameter('оп+рф').Set(op + ' ' + rf)


def set_sort(el):
    if not el.LookupParameter('Сортировка строка') and not el.LookupParameter('Сортировка систем'):
        print_once('Не найден параметр "Сортировка строка" ("Сортировка систем")')
        return
    sys_name = el.LookupParameter('ХТ Имя системы').AsString()
    if not sys_name:
        print_once('Ошибка сортировки. Не заполнен параметр "ХТ Имя системы"')
        return
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
    el.LookupParameter('Сортировка строка') and el.LookupParameter('Сортировка строка').Set(sort)
    el.LookupParameter('Сортировка систем') and el.LookupParameter('Сортировка систем').Set(sort)


def set_symbol(el):
    if el in flex_duct_curves:
        if el.LookupParameter('Классификация систем').AsString() == 'Приточный воздух' \
                and el.LookupParameter('Тип').AsValueString() != 'SG_Круглый изолированный':
            if 'symbol_flex_duct_curves' not in globals():
                symbols = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FlexDuctCurves).WhereElementIsElementType().ToElements()
                symbols = [el for el in symbols if el.LookupParameter('Имя типа').AsString() == 'SG_Круглый изолированный']
                symbol = symbols[0] if symbols else None
                globals()['symbol_flex_duct_curves'] = symbol
            if symbol_flex_duct_curves:
                el.FlexDuctType = symbol_flex_duct_curves
            else:
                print_once('Не найден тип гибкого воздуховода "SG_Круглый изолированный"')


def check_cost(el):
    cost = doc.GetElement(el.GetTypeId()).LookupParameter('Стоимость').HasValue
    if not cost:
        if 'bad_symbols' not in globals():
            globals()['bad_symbols'] = {}
        definintion = doc.GetElement(el.GetTypeId()).LookupParameter('Описание').AsString() or ''
        if 'Не учитывать' in definintion:
            return
        bad_symbols[el.GetTypeId()] = (output.linkify(el.Id), el.Id)


def result_check_cost():
    bad_symbols = globals().get('bad_symbols')
    if bad_symbols:
        bad_symbols_keys = bad_symbols.keys()
        bad_symbols_keys = sorted(bad_symbols_keys, key=lambda x: doc.GetElement(x).LookupParameter('Имя типа').AsString())
        bad_symbols_keys = sorted(bad_symbols_keys, key=lambda x: doc.GetElement(x).LookupParameter('Имя семейства').AsString())
        bad_symbols_keys = sorted(bad_symbols_keys, key=lambda x: doc.GetElement(x).LookupParameter('Категория').AsValueString())
        print('Следует прописать стоимость для сортировки и подсчёта количества:\n{}'.format(
                '\n'.join(['{} {}: {}: {}: {}'.format(
                    bad_symbols[k][0],
                    doc.GetElement(k).LookupParameter('Категория').AsValueString(),
                    doc.GetElement(k).LookupParameter('Имя типа').AsString(),
                    doc.GetElement(k).LookupParameter('Имя семейства').AsString(),
                    doc.GetElement(k).LookupParameter('Описание').AsString() or '',
                    ) for k in bad_symbols_keys])
                )
            )
    # el_ids = List[ElementId]([bad_symbols[key][1] for key in bad_symbols])
    # uidoc.Selection.SetElementIds(el_ids)


# -----------------------------------------------------------------------------------------
# -----------------------------------------------------------------------------------------
# -----------------------------------------------------------------------------------------

from pyrevit.forms import ProgressBar
with ProgressBar() as pb:
    t = Transaction(doc, 'Расчёт спеки')
    t.Start()
    try:
        len_elems = len(elems)
        print('N={}'.format(len_elems))
        counter = len_elems / 100
        for index, el in enumerate(elems):
            current = index / len_elems * 100
            if current > counter:
                pb.update_progress(current, 100)
                counter += len_elems / 100
            set_sys_name(el)       # Трубы, Фитинги и арматура труб, Оборудование и сантехника, Изоляция труб, Гибкие трубы, Воздуховоды, Гибкие воздуховоды, Изоляция воздуховодов, СДВиАВ, Воздухораспределители
            set_description(el)    # Трубы, Фитинги и арматура труб, Оборудование и сантехника, Изоляция труб, Гибкие трубы, Воздуховоды, Гибкие воздуховоды, Изоляция воздуховодов, СДВиАВ, Воздухораспределители
            set_mark(el)           # Трубы, Фитинги и арматура труб, Оборудование и сантехника, Изоляция труб, Гибкие трубы, Воздуховоды, Гибкие воздуховоды, Изоляция воздуховодов, СДВиАВ, Воздухораспределители
            set_amount(el)         # Трубы, Фитинги и арматура труб                             Изоляция труб, Гибкие трубы, Воздуховоды, Гибкие воздуховоды, Изоляция воздуховодов
            set_code(el)           #        Фитинги и арматура труб, Оборудование и сантехника
            set_subcomponents(el)  #                                 Оборудование
            set_op_rf(el)          # Трубы, Фитинги и арматура труб, Оборудование и сантехника, Изоляция труб, Гибкие трубы, Воздуховоды, Гибкие воздуховоды, Изоляция воздуховодов, СДВиАВ, Воздухораспределители
            set_symbol(el)         #                                                                                                      Гибкие воздуховоды
            set_sort(el)
            check_cost(el)

        result_check_cost()
    except Exception:
        t.RollBack()
        print(el)
        print('Ошибка {}, Δ={} с'.format(time.ctime().split()[3], time.time() - startTime))
        print(output.linkify(
            el.Id,
            '{}: {}: {}: id {}'.format(
                el.LookupParameter('Категория').AsValueString(),
                el.LookupParameter('Семейство').AsValueString(),
                el.LookupParameter('Тип').AsValueString(),
                el.Id,
            )
        ))
        output.show()
        raise
    t.SetName('Расчёт спеки {}, Δ={} с'.format(time.ctime().split()[3], time.time() - startTime))
    t.Commit()