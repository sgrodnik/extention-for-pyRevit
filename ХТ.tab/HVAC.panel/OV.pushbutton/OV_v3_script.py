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
kk = 10.76391041671
kkk = k * kk / 1000
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

elems = []
for old_var_name, var_name, bic in categories:
    globals()[var_name] = list(FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType().ToElements())
    elems.append(globals()[var_name])


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
    if el in plumbing_fixtures:
        sys_name = 'В1'
    else:
        sys_name = clean_sys_name(el.LookupParameter('Имя системы').AsString())
    el.LookupParameter('ХТ Имя системы').Set(sys_name)
    el.LookupParameter('ХТ Имя системы отступ').Set('‎' + ' ' * 90 + sys_name)


def set_description(el):
    symbol = doc.GetElement(el.GetTypeId())
    description = symbol.LookupParameter('Описание').AsString() or ''
    if not description:
        description = '--- Ошибка. Не заполнено Описание для {}: {}'.format(
            el.LookupParameter('Семейство').AsValueString(),
            el.LookupParameter('Тип').AsValueString(),
        )
    elif el in pipe_curves and 'егмент' in description:
        segment_description = el.LookupParameter('Описание сегмента').AsValueString()
        el.LookupParameter('ADSK_Наименование').Set(segment_description)
        return
    el.LookupParameter('ADSK_Наименование').Set(description)


def get_mark_pipe_curves(el):
    size = pipe_assortment.get(el.LookupParameter('Диаметр').AsDouble() * k)
    if not size:
        symbol_name = el.LookupParameter('Тип').AsValueString()
        segment_name = el.LookupParameter('Сегмент трубы').AsValueString()
        if '3262-75' in symbol_name:
            d_inner = el.LookupParameter('Внутренний диаметр').AsDouble() * k
            d_outer = el.LookupParameter('Внешний диаметр').AsDouble() * k
            thickness = (d_outer - d_inner) / 2.0
            size = 'ø{:n}×{:.1f}'.format(size, thickness).replace('.', ',')
        elif '10704-91' in symbol_name or 'ПВХ Dyka' in segment_name:
            d_inner = el.LookupParameter('Внутренний диаметр').AsDouble() * k
            d_outer = el.LookupParameter('Внешний диаметр').AsDouble() * k
            thickness = (d_outer - d_inner) / 2.0
            size = '∅{:n}×{:.1f}'.format(d_outer, thickness).replace('.', ',')
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


def get_mark_pipe_fitting_pipe_accessory(el):
    symbol_name = el.LookupParameter('Тип').AsValueString()
    if 'айка' in symbol_name or 'айба' in symbol_name:
        mark = 'М{:n}'.format(el.LookupParameter('d').AsDouble() * k)  # это зло надо перенести в каркас несущий
    else:
        size = el.LookupParameter('Размер').AsString().replace(',', '.')
        size_list = sorted([float(el) for el in set(size)], reverse=True) if size[0] else []
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
    100:'30×108 (для Ду100)',
    125:'30×133 (для Ду125)',
    150:'30×160 (для Ду150)',
    200:'30×219 (для Ду200)',
    250:'30×273 (для Ду250)',
}


def get_mark_pipe_insulations(el):
    host = doc.GetElement(el.HostElementId)
    if isinstance(host, Pipe):
        diameter = int(host.LookupParameter('Диаметр').AsDouble() * k)
        size = pipe_insulation_assortment.get(diameter)
        if not size:
            size = 'Ошибка: Не найден ø{}'.format(diameter)
            print('Для {} не найден диаметр изоляции: {}'.format(output.linkify(el.Id), diameter))
    else:
        diameter = int(host.LookupParameter('Размер').AsString().split('-')[0])
        length = diameter * 2.5 / k
        size = pipe_insulation_assortment.get(diameter, 'Ошибка: Не найден ø{}'.format(diameter))

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
        name = doc.GetElement(i.GetTypeId()).LookupParameter('Группа модели').AsString() or ''
    return name


def set_mark(el):
    if el in pipe_curves:
        mark = get_mark_pipe_curves(el)
    elif el in pipe_fitting + pipe_accessory:
        mark = get_mark_pipe_fitting_pipe_accessory(el)
    elif el in mechanical_equipment + plumbing_fixtures:
        mark = doc.GetElement(el.GetTypeId()).LookupParameter('Комментарии к типоразмеру').AsString()
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

    el.LookupParameter('ХТ Размер фитинга ОВ') and el.LookupParameter('ХТ Размер фитинга ОВ').Set(mark)
    el.LookupParameter('ADSK_Марка') and el.LookupParameter('ADSK_Марка').Set(mark)


def get_length_coef(el):
    symbol = doc.GetElement(el.GetTypeId())
    param = symbol.LookupParameter('Код по классификатору').AsString()
    coefficient = float(param) if param else 1.1
    return coefficient


def set_length(el):  # 2020.05.08 переименовать в set_amount
    coefficient = get_length_coef(el)

    if el in pipe_insulations:
        host = doc.GetElement(el.HostElementId)
        if isinstance(host, Pipe):
            length = el.LookupParameter('Длина').AsDouble() * k * coefficient
        else:
            diameter = int(host.LookupParameter('Размер').AsString().split('-')[0])
            length = diameter * 2.5 / k * coefficient

    elif el in duct_insulations:
        host = doc.GetElement(el.HostElementId)
        if host.LookupParameter('ХТ Площадь'):
            area = host.LookupParameter('ХТ Площадь').AsDouble() / kkk * coefficient
        elif el.LookupParameter('Площадь') and el.LookupParameter('Площадь').AsDouble() > 0:
            area = el.LookupParameter('Площадь').AsDouble() / kkk * coefficient
        else:
            raise Exception('Надо обработать этот вариант')
        length = area

    else:
        param = el.LookupParameter('Длина')
        length = param.AsDouble() * k * coefficient if param else 0

    el.LookupParameter('ХТ Длина ОВ') and el.LookupParameter('ХТ Длина ОВ').Set(length)
    el.LookupParameter('ADSK_Количество') and el.LookupParameter('ADSK_Количество').Set(length)


def set_code(el):
    code = doc.GetElement(el.GetTypeId()).LookupParameter('Код по классификатору').AsString() or ''
    if 'чтено в' in code:
        el.LookupParameter('ХТ Размер фитинга ОВ').Set('')
        el.LookupParameter('ADSK_Примечание') and el.LookupParameter('ADSK_Примечание').Set(code)
    else:
        code and code != '0' and el.LookupParameter('ADSK_Код изделия').Set(code)
    if el.LookupParameter('Артикул') and el.LookupParameter('ADSK_Код изделия'):
        el.LookupParameter('ADSK_Код изделия').Set(el.LookupParameter('Артикул').AsString())


def set_subcomponents(el):
    if hasattr(el, 'GetSubComponentIds'):
        if el.GetSubComponentIds():
            parent_sys_name = clean_sys_name(el.LookupParameter('ХТ Имя системы').AsString())
            for sub_el_id in el.GetSubComponentIds():
                doc.GetElement(sub_el_id).LookupParameter('ХТ Имя системы').Set(parent_sys_name)


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
# -----------------------------------------------------------------------------------------
# -----------------------------------------------------------------------------------------

for el in elems:
    set_sys_name(el)       # Трубы, Фитинги и арматура труб, Оборудование и сантехника, Изоляция труб, Гибкие трубы, Воздуховоды, Гибкие воздуховоды, Изоляция воздуховодов, СДВиАВ, Воздухораспределители
    set_description(el)    # Трубы, Фитинги и арматура труб, Оборудование и сантехника, Изоляция труб, Гибкие трубы, Воздуховоды, Гибкие воздуховоды, Изоляция воздуховодов, СДВиАВ, Воздухораспределители
    set_mark(el)           # Трубы, Фитинги и арматура труб, Оборудование и сантехника, Изоляция труб, Гибкие трубы, Воздуховоды, Гибкие воздуховоды, Изоляция воздуховодов, СДВиАВ, Воздухораспределители
    set_length(el)         # Трубы, Фитинги и арматура труб                             Изоляция труб, Гибкие трубы, Воздуховоды, Гибкие воздуховоды, Изоляция воздуховодов
    set_code(el)           #        Фитинги и арматура труб, Оборудование и сантехника
    set_subcomponents(el)  #                                 Оборудование
    set_op_rf(el)          # Трубы, Фитинги и арматура труб, Оборудование и сантехника, Изоляция труб, Гибкие трубы, Воздуховоды, Гибкие воздуховоды, Изоляция воздуховодов, СДВиАВ, Воздухораспределители
    set_symbol(el)         #                                                                                                      Гибкие воздуховоды
    set_sort(el)
    check_cost(el)

result_check_cost(el)

t.SetName('Расчёт спеки {}, Δ={} с'.format(time.ctime().split()[3], time.time() - startTime))
t.Commit()