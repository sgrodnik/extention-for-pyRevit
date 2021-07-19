# -*- coding: utf-8 -*-
""""""
from __future__ import print_function
__title__ = 'Test'
__author__ = 'SG'

import clr
import re
import os
import sys
import math

clr.AddReference('System.Core')
from System.Collections.Generic import *
from Autodesk.Revit.DB import SectionType, ElementId, PartUtils, ViewOrientation3D, XYZ, FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, Line, Structure, ParameterFilterElement, ParameterFilterRuleFactory, ElementParameterFilter, ViewDuplicateOption
from Autodesk.Revit.DB import WorksharingUtils, Structure
import Autodesk.Revit.DB as DB
import Autodesk.Revit.Exceptions as Exceptions
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from pyrevit import script
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application
output = script.get_output()
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


class CustomISelectionFilter(ISelectionFilter):
    def __init__(self, nom_categorie):
        self.nom_categorie = nom_categorie

    def AllowElement(self, e):
        if self.nom_categorie in e.Category.Name:
            return True
        else:
            return False

    def AllowReference(self, ref, point):
        return true

if __shiftclick__:
    os.startfile(os.path.realpath(__file__))
    sys.exit()


sel = [doc.GetElement(elid) for elid in uidoc.Selection.GetElementIds()]


t = Transaction(doc, 'Test')
t.Start()
# t.SetName('Test: Название')
# t.Commit()



t.SetName('Test: Список листов')
# data = []
# for el in FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsNotElementType().ToElements():
#     data.append([
#     doc.GetElement(el.OwnerViewId).LookupParameter('ADSK_Штамп Раздел проекта').AsString() or '',
#     doc.GetElement(el.OwnerViewId).LookupParameter('Номер листа').AsString(),
#     doc.GetElement(el.OwnerViewId).LookupParameter('Имя листа').AsString(),
#     el.LookupParameter('Тип').AsValueString().replace('А3А', 'А3А #').replace('Обложка и титул', 'А4К #'),
#     ])
# data = natural_sorted(data, key=lambda x: x[1])
# data = natural_sorted(data, key=lambda x: x[0])
# output.print_table(
#     table_data=data,
#     columns=['ADSK_Штамп Раздел проекта', 'Номер листа', 'Имя листа', 'Семейство и типоразмер', ],
# )
if [el for el in sel if el.LookupParameter('Категория').AsValueString() == 'Листы']:
    viewSet = DB.ViewSet()
    [viewSet.Insert(v) for v in sel]
    doc.PrintManager.PrintRange = DB.PrintRange.Select
    vss = doc.PrintManager.ViewSheetSetting  # если не работает, открой-закрой диалог печати вручную, напиши трай ексепт
    vss.CurrentViewSheetSet.Views = viewSet
    from datetime import datetime
    try:
        vss.SaveAs(datetime.today().strftime('! %Y.%m.%d %H:%M'))
    except:
        vss.SaveAs(datetime.today().strftime('! %Y.%m.%d %H:%M:%S'))

t.Commit()




# t.SetName('Test: ADSK_Наименование краткое из Описания')
# # for el in FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctCurves).WhereElementIsNotElementType().ToElements():
# for el in sel:
#     symbol = doc.GetElement(el.GetTypeId())
#     if symbol.LookupParameter('ADSK_Наименование краткое'):
#         try:
#             symbol.LookupParameter('ADSK_Наименование краткое').Set(symbol.LookupParameter('Описание').AsString() or '')
#         except:
#             pass
# t.Commit()






# t.SetName('Test: Название')
# from pyrevit import forms
# items = ['item1', 'item2', 'item3']
# res = forms.SelectFromList.show(items, button_name='Select Item')
# print(res)

# def main():
#     gl = globals()

#     print(*[gl[i].name for i in gl.keys() if hasattr(gl[i], 'name') and gl[i].name and 'Test: ' in gl[i].name])

#     t.Commit()






# def f123():
#     name = 'Test: Загиб выносок марок'
#     t.SetName(name)
#     for tag in sel:
#         tag.LeaderEnd = tag.LeaderElbow + XYZ(0, .65, 0)
#     t.Commit()
# # f123()




# def f12333():
#     name = 'Test: Загиб выносок марок333'
#     t.SetName(name)
#     for tag in sel:
#         tag.LeaderEnd = tag.LeaderElbow + XYZ(0, .65, 0)
#     t.Commit()
# # f123()






# t.SetName('Test: Часть левая или правая')
# for el in sel:
#     cir = doc.GetElement(ElementId(int(el.LookupParameter('Цепь').AsString())))
#     el.LookupParameter('Часть').Set(cir.LookupParameter('Часть').AsString()) if cir.LookupParameter('Часть').AsString() else None
# t.Commit()






# t.SetName('Test: Заполнение (копирование) обозначений цепей')
# els = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalCircuit).WhereElementIsNotElementType().ToElements())
# for el in els:
#     # name = el.LookupParameter('Цепь - Обозначение').AsString()
#     # el.LookupParameter('Номер связи').Set(name)
#     el.LookupParameter('Цепь - Обозначение').Set(str(el.Id.IntegerValue))
# t.Commit()




# t.SetName('Test: Кабельная трасса в Комментарии')
# els = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_CableTray).WhereElementIsNotElementType().ToElements())
# els += list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Conduit).WhereElementIsNotElementType().ToElements())
# els += list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ConduitFitting).WhereElementIsNotElementType().ToElements())
# subs = {
#     'F': ['FTP', 8],
#     'S': ['SDI', 6],
#     'U': ['UTP', 8],
#     'А': ['Аудио', 6],
#     'К': ['КПС', 6],
#     'О': ['ОПТ', 5],
# }
# for el in els:
#     if el.LookupParameter('Кабельная трасса'):
#         # res = el.LookupParameter('Кабельная трасса').AsString()
#         # res = ', '.join(res.split('\n'))
#         # el.LookupParameter('Комментарии').Set(res)
#         ids = el.LookupParameter('Кабельная трасса').AsString()
#         ids = ids.split('\n')
#         names = [el_.LookupParameter('Номер связи').AsString() for el_ in [doc.GetElement(ElementId(int(i))) for i in ids]]
#         names = ', '.join(names)
#         if el.LookupParameter('Категория').AsValueString() != 'Кабельные лотки':
#             el.LookupParameter('Комментарии').Set(names)
#             # continue
#         cirs = [doc.GetElement(ElementId(int(i))) for i in ids]
#         names_by_room = {}
#         prefix_count = {}
#         for cir in cirs:
#             room = cir.LookupParameter('Помещение').AsString()
#             name = cir.LookupParameter('Номер связи').AsString()
#             prefix = name[0]
#             if prefix not in prefix_count:
#                 prefix_count[prefix] = 0
#             prefix_count[prefix] += 1
#             # (prefix_count[prefix] +=1) if prefix in prefix_count else (prefix_count[prefix] = 0)
#             if room not in names_by_room:
#                 names_by_room[room] = []
#             names_by_room[room].append(name)
#         occupants = []
#         for room in natural_sorted(names_by_room.keys()):
#             if len(names_by_room.keys()) > 1:
#                 item = '{}: {}'.format(room, ', '.join(natural_sorted(names_by_room[room])))
#             else:
#                 item = ', '.join(natural_sorted(names_by_room[room]))
#             occupants.append(item)
#         # for prefix in prefix_count:
#         #     count = ', '.join(['{}x{}'.format(prefix, prefix_count[prefix])])
#         prefs = sorted(prefix_count)
#         prefs = sorted(prefs, key=lambda x: -prefix_count[x])
#         groups = [[subs.get(prefix, [prefix, 8]), prefix_count[prefix]] for prefix in prefs]
#         # groups = ['{}×{}'.format(subs.get(prefix, prefix), prefix_count[prefix]) for prefix in prefs]
#         counts = ' + '.join(['{}×{}'.format(i[0], n) if n > 1 else i[0] for i, n in groups])
#         formula = ' + '.join(['{}²×{}'.format(i[1], n) if n > 1 else '{}²'.format(i[1]) for i, n in groups])
#         area_sum = sum([i[1] ** 2 * n for i, n in groups])
#         formula += ' = {} (мм²)'.format(area_sum)
#         if el.LookupParameter('Ширина'):
#             area = el.LookupParameter('Ширина').AsDouble() * k * el.LookupParameter('Высота').AsDouble() * k
#             formula += '\nЗаполнение {} / {:.0f} = {:.0f} %.'.format(area_sum, area, area_sum / area * 100)
#         # for k, v in subs.items():
#         #     counts = re.compile(re.escape(k), re.IGNORECASE).sub(v, counts)
#         occupants = ';\n'.join(occupants) + '.'
#         if el.LookupParameter('Кем занято'):
#             el.LookupParameter('Кем занято').Set(counts + ':\n' + occupants)
#         if el.LookupParameter('Заполнение'):
#             el.LookupParameter('Заполнение').Set(formula)
#             el.LookupParameter('Заполнение %').Set(area_sum / area * 100)
#         if el.LookupParameter('Категория').AsValueString() == 'Кабельные лотки':
#             el.LookupParameter('Комментарии').Set(counts.replace(' + ', ', '))

#         # сравнить area с сортаментом
#         # import json
#         # symbol = doc.GetElement(el.GetTypeId())
#         # if not symbol.LookupParameter('Сортамент'):
#         #     continue
#         # param = symbol.LookupParameter('Сортамент').AsString()
#         # settings = json.loads(param or '{}')
#         # if not settings:
#         #     continue
#         # assort = settings['ADSK_Наименование'].keys()
#         # assort = [i.replace('h', '').split('x') for i in assort]

# # Помещения и участок
# els = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_CableTray).WhereElementIsNotElementType().ToElements())
# # els = natural_sorted(els, key=lambda x: x.LookupParameter('Размер').AsString())
# # els = natural_sorted(els, key=lambda x: x.LookupParameter('Комментарии').AsString())
# dct1 = {}
# dct2 = {}
# i = 1
# for el in els:
#     key = el.LookupParameter('Кем занято').AsString() + el.LookupParameter('Размер').AsString()
#     # key = key.replace(' ', '').replace('\n', '')
#     if key not in dct1:
#         dct1[key] = []
#     dct1[key].append(el.LookupParameter('Помещение').AsString() if el.LookupParameter('Помещение').AsString() else '-')

#     key = el.LookupParameter('Комментарии').AsString() + el.LookupParameter('Размер').AsString()
#     # key = key.replace(' ', '').replace('\n', '')
#     if key not in dct2:
#         dct2[key] = i
#         i += 1
# for el in els:
#     key = el.LookupParameter('Кем занято').AsString() + el.LookupParameter('Размер').AsString()
#     # key = key.replace(' ', '').replace('\n', '')
#     res = list(set(dct1[key]))
#     res = natural_sorted(res)
#     el.LookupParameter('Помещения').Set(', '.join(res))

#     key = el.LookupParameter('Комментарии').AsString() + el.LookupParameter('Размер').AsString()
#     el.LookupParameter('Состав').Set(str(dct2[key]))

# t.Commit()






# t.SetName('Test: Выбор элементров по маркам')

# uidoc.Selection.SetElementIds(List[ElementId]([el.TaggedLocalElementId for el in sel]))

# t.Commit()








# t.SetName('Test: Рисование логарифмической шкалы')

# # P1 = XYZ(x1,y1,z1)
# # P2 = XYZ(x2,y2,z2)
# # L1 = Line.CreateBound(P1,P2)
# # doc.Create.NewDetailCurve(vd,L1)

# # lst = [i + 1 for i in range(9)] + [(i + 1) * 10 for i in range(9)] + [(i + 1) * 100 for i in range(10)]
# lst = [.01, .02, .03, .04, .05, .06, .07, .08, .09,
#        .1, .2, .3, .4, .5, .6, .7, .8, .9,
#        1, 2, 3, 4, 5, 6, 7, 8, 9,
#        10, 20, 30, 40, 50, 60, 70, 80, 90,
#        100, 200, 300, 400, 500, 600, 700, 800, 900, 1000, ]
# # lst = [math.log10(i) * 1 for i in lst]


# for n in lst:
#     # x = math.log10(n) * .614
#     x = math.log10(n) * 1
#     x = x / k * 1000
#     newline = DB.Line.CreateBound(XYZ(0, x, 0), XYZ(100/k, x, 0))
#     doc.Create.NewDetailCurve(doc.ActiveView, newline)
#     tn = DB.TextNote.Create(doc, doc.ActiveView.Id, XYZ(0, x, 0), str(n), ElementId(570635))
#     # tn.UpDirection = XYZ(1, 0, 0)

# t.Commit()







# import clr
# clr.AddReference('System.Windows.Forms')
# clr.AddReference('IronPython.Wpf')
# from pyrevit import UI
# from pyrevit import script
# xamlfile = script.get_bundle_file('test.xaml')

# import wpf
# from System import Windows

# class MyWindow(Windows.Window):
#     def __init__(self):
#         wpf.LoadComponent(self, xamlfile)

#     # @property
#     # def user_name(self):
#     #     return self.textbox.Text

#     # def say_hello(self, sender, args):
#     #     UI.TaskDialog.Show(
#     #         "Hello World",
#     #         "Hello {}".format(self.user_name or 'World')
#     #         )

# MyWindow().ShowDialog()







# t.SetName('Test: просмотр параметров')
# dic = {}
# for el in sel:
#     dic[el.Id.IntegerValue] = {par.Definition.Name: par.AsValueString() or par.AsString() for par in sorted(el.Parameters, key=lambda x: x.Definition.Name)}

# parnames = set()
# for num in dic:
#     for parname in dic[num]:
#         parnames.add(parname)
# parnames = sorted(list(parnames))
# # print('Ids:%%%{}'.format('%%%'.join([str(num) for num in dic])))
# # for name in parnames:
# #     print('{}%%%{}'.format(name, '%%%'.join([dic[num][name] or '---' for num in dic])))

# data = [[name] + [dic[num][name] or '---' for num in dic] for name in parnames]

# output.print_table(
#     table_data=data,
#     columns=['Ids'] + [str(num) for num in dic],
#     formats=['', '', '', ],
#     # last_line_style='color:red;'
# )

# '''
# 123456: [
#         'parname1', 'val',
#         'parname2', 'val',
#         'parname3', 'val',
#         ],
# 123457: [
#         'parname1', 'val',
#         'parname2', 'val',
#         'parname3', 'val',
#         ],

# data = [
#     ['row1', 'data', 'data', 80 ],
#     ['row2', 'data', 'data', 45 ],
#     ]

# '''
# t.Commit()





# t.SetName('Test: заполнение параметров')

# els = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsElementType().ToElements())
# src = {
#     'Семейство+++Тип': ('КабельСигнальный', 'КабельСиловой', 'ХТ УГО'),
#     'В стойке дополнительное+++Кабельный органайзер 2U': ('', '', 'Кабельный органайзер 2U'),
#     'В стойке дополнительное+++Кабельный органайзер боковой': ('', '', 'Кабельный органайзер боковой'),
#     'В стойке дополнительное+++Комплект щёточного ввода в шкаф 2': ('', '', 'Комплект щёточного ввода в шкаф'),
#     'Компьютер+++Компьютер': ('', '', ''),
#     'Свитч+++Матричный коммутатор SDI': ('utpp', 'По кратчайшему пути', 'Матричный коммутатор SDI'),
#     'В стойке дополнительное+++Невыдвижная полка': ('', '', 'Полка'),
#     'В стойке дополнительное+++Патч-панель 24': ('', '', 'Патч-панель 24×RJ-45'),
#     'В стойке дополнительное+++Розеточный блок': ('', '', 'Розеточный блок'),
#     'Фейк+++Фейк 1': ('', '', ''),
#     'Фейк+++Фейк 2': ('', '', ''),
#     'Фейк+++Фейк 3': ('', '', ''),
#     'Фейк+++Фейк 4': ('', '', ''),
#     'Фейк+++Фейк 5': ('', '', ''),
#     'Фейк+++Фейк 6': ('', '', ''),
#     'Фейк+++Фейк 7': ('', '', ''),
#     'Фейк+++Фейк 8': ('', '', ''),
#     'Фейк+++Фейк 9': ('', '', ''),
#     'Фейк+++Фейк 10': ('', '', ''),
#     'Фейк+++Фейк 11': ('', '', ''),
#     'Фейк+++Фейк 12': ('', '', ''),
#     'Фейк+++Фейк 13': ('', '', ''),
#     'Стойка серверная+++Шкаф 600×800 37U': ('', '', 'Шкаф 37U'),
#     'В стойке+++Server HP': ('utpp', 'По кратчайшему пути', 'Сервер Hospital platform'),
#     'В стойке+++Server OR патчкорд': ('utpp', 'По кратчайшему пути', 'Сервер операционной'),
#     'В стойке+++Блок B4': ('sdi/sdi/sdi/sdi/utpp/aud', 'По кратчайшему пути', 'Блок захвата и хранения видеопотоков B4'),
#     'В стойке+++ИБП': ('utpp', 'По кратчайшему пути', 'ИБП'),
#     'В стойке+++Коммутатор сетевой LAN HP': ('utp', 'По кратчайшему пути', 'Сетевой коммутатор госпитальной платформы (LAN HP)'),
#     'В стойке+++Коммутатор сетевой LAN OR': ('utp', 'По кратчайшему пути', 'Сетевой коммутатор операционной (LAN OR)'),
#     'В стойке+++Конвертер HDMI to SDI для 49"': ('sdi', 'По кратчайшему пути', 'HDMI>SDI'),
#     'В стойке+++Конвертер HDMI to SDI для конференц-связи': ('sdi', 'По кратчайшему пути', 'HDMI>SDI'),
#     'В стойке+++Приёмник FTP to Audio для радиомикрофона': ('audio', 'По кратчайшему пути', 'FTP>Audio'),
#     'В стойке+++Сплиттер': ('aud/aud', 'По кратчайшему пути', 'Сплиттер'),
#     'В стойке+++Сумматор': ('aud', 'По кратчайшему пути', 'Сумматор'),
#     'В стойке+++Усилитель': ('aud', 'По кратчайшему пути', 'Усилитель'),
#     'Вывод кабеля+++Видеоисточник SDI': ('sdi', 'По кратчайшему пути', 'SDI'),
#     'Конвертеры+++Конвертер HDMI to SDI для ГМУ 32': ('sdi', 'По кратчайшему пути', ''),
#     'Конвертеры+++Конвертер SDI to HDMI для 49"': ('sdi', 'По кратчайшему пути', ''),
#     'Конвертеры+++Конвертер SDI to HDMI для ГМУ 32': ('sdi', 'По кратчайшему пути', ''),
#     'Конвертеры+++Конвертер YPbPr to 3G-SDI': ('sdi', 'По кратчайшему пути', ''),
#     'Конвертеры+++Передатчик Audio to FTP для радиомикрофона': ('ftp', 'По кратчайшему пути', ''),
#     'Щиты+++Щит автоматики ОВ': ('utp/ftp/pwr', 'По кратчайшему пути', 'ЩАОВ'),
#     'Щиты+++Щит силовой': ('pwr/pwr/pwr/pwr/utp/utp/utp', '', 'ЩС'),
#     'Камера+++Cветильник хирургический с камерой': ('sdi/utp', 'По кратчайшему пути', ''),
#     'Камера+++Камера обзорная': ('sdi/utp', 'По кратчайшему пути', ''),
#     'БААВС+++БААВС': ('sdi/sdi', 'По кратчайшему пути', ''),
#     'Инженерка+++Блок управления приводом откатной двери': ('utp', 'через верх', ''),
#     'Инженерка+++Блок управления приводом распашной двери': ('utp', 'через верх', ''),
#     'Инженерка+++Блок управления приводом жалюзи': ('pwr', 'через верх', ''),
#     'Инженерка+++Трассировка': ('', 'По кратчайшему пути', ''),
#     'Инженерка+++Датчик перепада давления': ('pwr', 'По кратчайшему пути', ''),
#     'Инженерка+++Датчик температуры и влажности': ('ftp', 'через верх', ''),
#     'Инженерка+++Динамик потолочный': ('audio', 'По кратчайшему пути', ''),
#     'Инженерка+++Радиомикрофон': ('aud', 'По кратчайшему пути', ''),
#     'Инженерка+++Кнопка ISERV': ('utp', 'По кратчайшему пути', ''),
#     'В стойке+++Миникомпьютер для 49"': ('utpp/audio', 'По кратчайшему пути', 'MiniPC 49"'),
#     'В стойке+++Миникомпьютер для ГМУ 32': ('utp/aud', 'По кратчайшему пути', 'MiniPC'),
#     'В стойке+++Миникомпьютер для конференц-связи': ('utpp/aud', 'По кратчайшему пути', 'MiniPC конф.'),
#     'В стойке+++Миникомпьютер для станции м/с': ('utp/aud', 'По кратчайшему пути', 'MiniPC'),
#     'Дали+++Блок контроля освещения (DALI)': ('pwr', 'По кратчайшему пути', ''),
#     'КОК+++Контрольно-отключающая коробка МГС (КОК)': ('utp/utp', 'По кратчайшему пути', ''),
#     'Мониторы и другие кубы+++Монитор обзорный 49"': ('hdmi((1))', 'По кратчайшему пути', '49"'),
#     'Мониторы и другие кубы+++Монитор управления 32" главный': ('((1))aud', 'По кратчайшему пути', '32"'),
#     'Мониторы и другие кубы+++Рабочая станция 27" с клавиатурой': ('utp', 'По кратчайшему пути', '27"'),
#     'Мониторы и другие кубы+++Блок управления инженерными системами': ('utpp/utp/utp/utp/utp', 'По кратчайшему пути', 'БУИС'),
#     'Розетка+++Розетка HDMI': ('hdmi', 'По кратчайшему пути', 'HDMI'),
# }
# for el in els:
#     name = el.LookupParameter('Имя семейства').AsString() + '+++' + el.LookupParameter('Имя типа').AsString()
#     # print(name)
#     if name in src:
#         print(name)
#         el.LookupParameter('КабельСигнальный').Set(src[name][0])
#         el.LookupParameter('КабельСиловой').Set(src[name][1])
#         el.LookupParameter('ХТ УГО').Set(src[name][2])

# t.Commit()













# t.SetName('Test: json')
# import json
# els = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_CableTray).WhereElementIsNotElementType().ToElements())
# els += list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Conduit).WhereElementIsNotElementType().ToElements())
# els += list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ConduitFitting).WhereElementIsNotElementType().ToElements())

# settings = {
#     'ADSK_Наименование': {16: 'Труба ПП не распространяющая горение гибкая гофр. д.16мм, лёгкая с протяжкой, 100м, цвет синий',
#                           20: 'Труба ПП не распространяющая горение гибкая гофр. д.20мм, лёгкая с протяжкой, 100м, цвет синий',
#                           25: 'Труба ПП не распространяющая горение гибкая гофр. д.25мм, лёгкая с протяжкой, 100м, цвет синий',
#                           32: 'Труба ПП не распространяющая горение гибкая гофр. д.32мм, лёгкая с протяжкой, 100м, цвет синий',
#                           },
#     'ADSK_Марка': {16: 'Octopus',
#                    20: 'Octopus',
#                    25: 'Octopus',
#                    32: 'Octopus',
#                    },
#     'ADSK_Код изделия': {16: '11916',
#                          20: '11920',
#                          25: '11925',
#                          32: '11932',
#                          },
# }
# settings = json.dumps(settings, sort_keys=True, indent=4, separators=(',', ': '))
# symbol = doc.GetElement(sel[0].GetTypeId())
# symbol.LookupParameter('Сортамент').Set(settings)
# # for el in els:
# #     el.LookupParameter('ADSK_Наименование').Set()
# #     el.LookupParameter('ADSK_Марка').Set()
# #     el.LookupParameter('ADSK_Код изделия').Set()

# t.Commit()




# t.SetName('Test: показать цепи')

# system_ids_to_select = []
# for el in sel:
#     # print(el.Id)
#     if el.MEPModel.ElectricalSystems:
#         for system in el.MEPModel.ElectricalSystems:
#             if system.Id not in system_ids_to_select:
#                 system_ids_to_select.append(system.Id)
#     elif 'ейк' in el.LookupParameter('Тип').AsValueString():
#         target = el.LookupParameter('Цепь').AsString()
#         system_ids_to_select.append(ElementId(int(target)))

# for id in system_ids_to_select:
#     el = doc.GetElement(id)
#     print('{} Пом. {}: {} {} – {} {} {}'.format(
#         output.linkify(id),
#         # el.LookupParameter('Помещение цепи').AsString() if el.LookupParameter('Помещение цепи') else None,
#         el.LookupParameter('Помещение цепи').AsString() if el.LookupParameter('Помещение цепи') else el.LookupParameter('Помещение').AsString() if el.LookupParameter('Помещение') else '',
#         el.LookupParameter('Позиция начала').AsString() if el.LookupParameter('Позиция начала') else '',
#         el.LookupParameter('Имя нагрузки').AsString() if el.LookupParameter('Имя нагрузки') else '',
#         el.LookupParameter('Позиция конца').AsString() if el.LookupParameter('Позиция конца') else '',
#         el.LookupParameter('Панель').AsString(),
#         el.LookupParameter('Тип, марка').AsString(),
#     ))
# if len(system_ids_to_select) > 1:
#     print('Выбрать все {}'.format(output.linkify(system_ids_to_select)))

# t.Commit()




                    # print('{} цепь {} не имеет панели'.format(cir_number, output.linkify(el.Id, name)))







# t.SetName('Test: построение линий')

# # line = sel[0]
# # print(line.GeometryCurve.Origin)
# # print(line.GeometryCurve.Direction)
# # newline = DB.Line.CreateBound(XYZ(0, 0, 0), XYZ(1, 0, 0))
# # line.GeometryCurve = newline
# # print(line.GeometryCurve.Origin)
# # print(line.GeometryCurve.Direction)


# def get_intersection(first_line, second_line):
#     if not isinstance(first_line, DB.Curve):
#         first_line = first_line.GeometryCurve
#     if not isinstance(second_line, DB.Curve):
#         second_line = second_line.GeometryCurve
#     results = clr.StrongBox[DB.IntersectionResultArray]()
#     first_line.Intersect(second_line, results)
#     enumerator = results.GetEnumerator()  # Если ошибка, вероятно нет пересечений и надо делать проверку
#     enumerator.MoveNext()
#     # intersection = enumerator.Current.XYZPoint
#     return enumerator.Current.XYZPoint


# def get_project(line, point):
#     result = line.GeometryCurve.Project(point)
#     # enumerator = results.GetEnumerator()
#     # enumerator.MoveNext()
#     # return enumerator.Current.XYZPoint
#     return result.XYZPoint


# def fillet_arc(first_line, second_line, radius=.7):
#     # Create a new arc defined by its center, radius, angles and 2 axes
#     center = get_intersection(first_line, second_line)
#     xAxis = XYZ(1, 0, 0)  # The x axis to define the arc plane. Must be normalized
#     yAxis = XYZ(0, 1, 0)  # The y axis to define the arc plane. Must be normalized
#     circle1 = DB.Arc.Create(center, radius, 0, 2 * math.pi, xAxis, yAxis)
#     # doc.Create.NewDetailCurve(doc.ActiveView, circle1)

#     inter1 = get_intersection(circle1, first_line)
#     circle2 = DB.Arc.Create(inter1, .05, 0, 2 * math.pi, xAxis, yAxis)
#     # doc.Create.NewDetailCurve(doc.ActiveView, circle2)

#     inter2 = get_intersection(circle1, second_line)
#     circle3 = DB.Arc.Create(inter2, .05, 0, 2 * math.pi, xAxis, yAxis)
#     # doc.Create.NewDetailCurve(doc.ActiveView, circle3)

#     ### chamfer = DB.Line.CreateBound(inter1, inter2)
#     ### doc.Create.NewDetailCurve(doc.ActiveView, chamfer)

#     # new_line = first_line.GeometryCurve.Clone()
#     # print(dir(new_line))
#     # myVector = XYZ(1, 0, 0)
#     tf = DB.Transform.CreateTranslation(inter2 - center)
#     first_line_offset = first_line.GeometryCurve.CreateTransformed(tf)
#     # doc.Create.NewDetailCurve(doc.ActiveView, first_line_offset)
#     tf = DB.Transform.CreateTranslation(inter1 - center)
#     second_line_offset = second_line.GeometryCurve.CreateTransformed(tf)
#     # doc.Create.NewDetailCurve(doc.ActiveView, second_line_offset)

#     inter3 = get_intersection(first_line_offset, second_line_offset)
#     circle4 = DB.Arc.Create(inter3, .05, 0, 2 * math.pi, xAxis, yAxis)
#     doc.Create.NewDetailCurve(doc.ActiveView, circle4)

#     project1 = get_project(first_line, inter3)
#     circle5 = DB.Arc.Create(project1, .05, 0, 2 * math.pi, xAxis, yAxis)
#     doc.Create.NewDetailCurve(doc.ActiveView, circle5)
#     project2 = get_project(second_line, inter3)
#     circle6 = DB.Arc.Create(project2, .05, 0, 2 * math.pi, xAxis, yAxis)
#     doc.Create.NewDetailCurve(doc.ActiveView, circle6)


#     # xAxis = (inter3 - project2).Normalize()  # The x axis to define the arc plane. Must be normalized
#     newline = DB.Line.CreateBound(XYZ(0, 0, 0), XYZ(0, .1, 0))
#     doc.Create.NewDetailCurve(doc.ActiveView, newline)
#     newline = DB.Line.CreateBound(XYZ(.1, 0, 0), XYZ(0, 0, 0))
#     doc.Create.NewDetailCurve(doc.ActiveView, newline)
#     newline = DB.Line.CreateBound(XYZ(0, 0, 0), project2 - inter1)
#     doc.Create.NewDetailCurve(doc.ActiveView, newline)


#     # start_angle = get_azimuth()
#     # end_angle = xAxis.AngleTo(project1)
#     # arc = DB.Arc.Create(inter3, (inter3 - project2).GetLength(), start_angle, end_angle, xAxis, yAxis)
#     # doc.Create.NewDetailCurve(doc.ActiveView, arc)


# first_line = doc.GetElement(ElementId(2661))
# second_line = doc.GetElement(ElementId(2483))
# fillet_arc(first_line, second_line)


# t.Commit()










# t.SetName('Test: Нумерация ADSK_Позиция типов аннотаций')
# if len(sel) == 1:
#     category_name = sel[0].LookupParameter('Категория').AsValueString()
#     symbol = doc.GetElement(sel[0].GetTypeId())
#     value = int(symbol.LookupParameter('ADSK_Позиция').AsString())
#     while True:
#         if not __shiftclick__:
#             value += 1
#         try:
#             target = uidoc.Selection.PickObject(ObjectType.Element, CustomISelectionFilter(category_name), 'Выберите следующий элемент ' + category_name + '" [ESC для выхода]')
#         except Exceptions.OperationCanceledException:
#             t.Commit()
#             break
#             # sys.exit()

#         target = doc.GetElement(target.ElementId)
#         target_symbol = doc.GetElement(target.GetTypeId())
#         target_symbol.LookupParameter('ADSK_Позиция').Set(str(value))
#         doc.Regenerate()








# t.Start()
# t.SetName('Test: Комментарии изоляции')
# duct_insulations = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctInsulations).WhereElementIsNotElementType().ToElements()
# for el in duct_insulations:
#     host = doc.GetElement(el.HostElementId)
#     if host.LookupParameter('Комментарии').AsString():  #  and host.LookupParameter('Комментарии').AsString() == "Антей":
#         el.LookupParameter('Комментарии').Set(host.LookupParameter('Комментарии').AsString())

# t.Commit()










# t.SetName('Test: Спецификация ручная импорт (Москва ИКБ ТМ)')

# path = os.path.join(os.path.dirname(doc.PathName), 'С ручная.txt')
# file_started = False
# try:
#     src = open(path, 'r').read().decode("utf-8")
#     if src:
#         src = src[:-1] if src[-1] == '\n' else src
#     if not src or src.startswith('\nЭтот'):
#         raise IOError
# except IOError:
#     info = '\nЭтот файл должен содержать таблицу (разделители - табуляция).\nПервая строка должна содержать заголовки. Последняя строка может быть пустой.\n' + path
#     with open(path, 'w') as file:
#         file.write(info.encode("utf-8"))
#     os.startfile(path)
#     src = '\n\n'
#     file_started = True

# if __shiftclick__:
#     if not file_started:
#         os.startfile(path)
#     sys.exit()

# if src and src != '\n\n':
#     titles = src.split('\n')[0].split('\t')
#     _pozitsiia = titles.index('Позиция')
#     _naimenovanie = titles.index('Наименование и техническая характеристика')
#     _tip = titles.index('Тип, марка, обозначение документа, опросного листа')
#     _kod = titles.index('Код оборудования, изделия, материала')
#     _izgotovitel = titles.index('Изготовитель')
#     _edinitsa = titles.index('Единица измерения')
#     _kolichestvo = titles.index('Количество')
#     _primechanie = titles.index('Примечание')

#     raws = [x.split('\t') for x in src.split('\n')][1:]
#     pozitsiia = [x[_pozitsiia] for x in raws]
#     naimenovanie = [x[_naimenovanie] for x in raws]
#     tip = [x[_tip] for x in raws]
#     kod = [x[_kod] for x in raws]
#     izgotovitel = [x[_izgotovitel] for x in raws]
#     edinitsa = [x[_edinitsa] for x in raws]
#     kolichestvo = [x[_kolichestvo] for x in raws]
#     primechanie = [x[_primechanie] for x in raws]

#     schedules = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Schedules).WhereElementIsNotElementType().ToElements()
#     schedule = list(filter(lambda x: x.Name == 'С ручная', schedules))[0]
#     data = schedule.GetTableData().GetSectionData(SectionType.Body)
#     for i in range(data.LastRowNumber): data.RemoveRow(i + 1)
#     for _ in pozitsiia: data.InsertRow(0)
#     els = FilteredElementCollector(doc, schedule.Id).ToElements()
#     for index, el in enumerate(els):
#         el.LookupParameter('к1').Set(pozitsiia[index])
#         el.LookupParameter('к2').Set(naimenovanie[index])
#         el.LookupParameter('к3').Set(tip[index])
#         el.LookupParameter('к4').Set(kod[index])
#         el.LookupParameter('к5').Set(izgotovitel[index])
#         el.LookupParameter('к6').Set(edinitsa[index])
#         el.LookupParameter('к7').Set(kolichestvo[index])
#         el.LookupParameter('к8').Set(primechanie[index])

# t.Commit()












# t.SetName('Test: Радиусы гофры')
# if len(sel) == 1:
#     category_name = sel[0].LookupParameter('Категория').AsValueString()
#     radius = sel[0].LookupParameter('Радиус загиба').AsDouble()
#     while True:
#         radius += 26 / k
#         try:
#             target = uidoc.Selection.PickObject(ObjectType.Element, CustomISelectionFilter(category_name), 'Выберите следующий элемент ' + category_name + '" [ESC для выхода]')
#         except Exceptions.OperationCanceledException:
#             t.Commit()
#             break
#             # sys.exit()

#         target = doc.GetElement(target.ElementId)
#         target.LookupParameter('Радиус загиба').Set(radius)
#         doc.Regenerate()



# t.SetName('Test: Распространение комментариев')
# comms = []
# engineering_flags = []
# for el in sel:
#     engineering_flag = el.LookupParameter('Инженерка').AsInteger()
#     engineering_flags.append(engineering_flag) if engineering_flag == 1 else None
#     comment = el.LookupParameter('Комментарии').AsString()
#     comment = comment if comment else ''
#     comms.append(comment) if comment else None
# comms = list(set(comms))
# engineering_flags = list(set(engineering_flags))
# if len(engineering_flags) == 1:
#     for el in sel:
#         el.LookupParameter('Инженерка').Set(engineering_flags[0])
# # else:
# #     print(len(engineering_flags))
# #     print(engineering_flags)
# #     print('--------------')
# if len(comms) == 1:
#     for el in sel:
#         el.LookupParameter('Комментарии').Set(comms[0])
# else:
#     print(len(comms))
#     print(comms)
# t.Commit()




# t.Start()
# t.SetName('Test: Сбор комментариев')
# for el in sel:
#     comment = el.LookupParameter('Комментарии').AsString()
#     comment = comment if comment else ''
#     for sub in comment.split(', '):
#         if '×' in sub:
#             sub, count = sub.split('×')
#         else:
#             sub, count = [sub, 1]
#         print(sub + '---' + str(count))
# t.Commit()









# t.Start()
# done = []
# for el in sel:
#   symbol = doc.GetElement(el.GetTypeId())
#   if s
# symbols = [doc.GetElement(el.GetTypeId()) for el in sel]
# symbols = [i for i in symbols if len()]
# for symbol in symbols:
#     val = symbol.LookupParameter('ADSK_Позиция').AsString()
#     val = int(val)
#     symbol.LookupParameter('ADSK_Позиция').Set(str(val + 1))
#     # cost = symbol.LookupParameter('Стоимость').AsDouble()
#     # symbol.LookupParameter('Стоимость').Set(cost + 1)

# t.Commit()









# sketch_planes = FilteredElementCollector(doc).OfClass(DB.SketchPlane).ToElements()
# print(len(sketch_planes))
# sketch_planes = [i for i in sketch_planes if i.OwnerViewId.IntegerValue == -1]
# print(len(sketch_planes))
# [doc.Delete(i.Id) for i in sketch_planes]
# print(len(sketch_planes))

# t.Commit()









# src = {'2.07.1.004': '1', '2.07.1.008': '2', '2.07.1.018': '3', '2.07.1.029': '4', '2.07.2.001': '5', '2.07.2.002': '6', '2.07.2.003': '7', '2.07.2.004': '8', '2.07.2.005': '9', '2.07.2.007': '10', '2.07.2.008': '11', '2.07.2.009': '12', '2.07.2.012': '13', '2.07.2.015': '15', '2.07.3.001': '16', '2.07.3.002': '17', '2.07.3.003': '18', '2.07.3.005': '19', '2.07.3.006': '20', '2.07.3.007': '21', '2.07.3.008': '22', '2.07.3.011': '24', '2.07.3.013': '25', '2.07.3.021': '26', '2.07.3.022': '27', '2.07.3.025': '29', '2.07.3.027': '30', '2.07.3.028': '31', '2.07.3.030': '33', '2.07.3.032': '34', '2.07.3.033': '35', '2.07.3.034': '36', '2.07.3.038': '37', '2.07.3.039': '38', '2.07.3.045': '39', '2.07.3.053': '40', '2.07.1.004': '1', '2.07.1.008': '2', '2.07.1.018': '3', '2.07.2.001': '5', '2.07.2.002': '6', '2.07.2.003': '7', '2.07.2.004': '8', '2.07.2.005': '9', '2.07.2.007': '10', '2.07.2.008': '11', '2.07.2.009': '12', '2.07.2.013': '14', '2.07.2.015': '15', '2.07.3.001': '16', '2.07.3.002': '17', '2.07.3.003': '18', '2.07.3.005': '19', '2.07.3.006': '20', '2.07.3.007': '21', '2.07.3.008': '22', '2.07.3.010': '23', '2.07.3.013': '25', '2.07.3.021': '26', '2.07.3.022': '27', '2.07.3.024': '28', '2.07.3.027': '30', '2.07.3.028': '31', '2.07.3.029': '32', '2.07.3.032': '34', '2.07.3.033': '35', '2.07.3.034': '36', '2.07.3.038': '37', '2.07.3.039': '38',}

# for el in sel:
#     symbol = doc.GetElement(el.GetTypeId())
#     comment = symbol.LookupParameter('Комментарии к типоразмеру').AsString()
#     # symbol.LookupParameter('ADSK_Позиция').Set(src[comment])
#     symbol.LookupParameter('Стоимость').Set(float(src[comment]))
#     # position = symbol.LookupParameter('ADSK_Позиция').AsString()


# t.Commit()


















# src = [
#     ('17', 'Epp 3122 000.051', 'Дозатор механический переменного объёма восьмиканальный Eppendorf Research Plus  30-300  мкл', '3'),
#     ('18', 'Epp 3122 000.213', 'Дозатор механический переменного объёма восьмиканальный Eppendorf Research Plus  120-1200  мкл', '3'),
#     ('19', '4641062', '1-КАНАЛЬНЫЙ ДОЗАТОР ТЕХНО F1* 2-20 мкл ДПОП-1-2-20, Thermo Fisher Scientific', '6'),
#     ('20', '4641082', '1-КАНАЛЬНЫЙ ДОЗАТОР ТЕХНО F1* 20-200 мкл ДПОП-1-20-200, Thermo Fisher Scientific', '6'),
#     ('21', '4641102', '1-КАНАЛЬНЫЙ ДОЗАТОР ТЕХНО F1* 100-1000 мкл ДПОП-1-100-1000, Thermo Fisher Scientific', '6'),
#     ('22', '4641112', '1-КАНАЛЬНЫЙ ДОЗАТОР ТЕХНО F1* 0.5-5 мл ДПОП-1-500-5000, Thermo Fisher Scientific', '6'),
#     ('23', 'LS-2', 'Штатив-подставка для пипеток универсальный на 5 дозаторов', '3'),
#     ('24', 'RA-5015', 'Штатив “рабочее место” для микропробирок 1.5 мл, 50 лунок', '10'),
#     ('25', 'RA-20002', 'Штатив "рабочее место" для микропробирок 0,2 мл, 200 лунок', '4'),
#     ('26', 'RA-5005', 'Штатив “рабочее место” для микропробирок 0.5 мл, 50 лунок', '10'),
#     ('27', 'SSI-1260-00', 'Пробирки типа Эппендорф, объёмом 1,5 мл, градуир., 500 шт/уп', '60'),
#     ('28', 'SSI-4337NSFS', 'Наконечники Vertex на 1000 мкл, с фильтром, бесцветные, градуированные, стерильные, 96 шт/штатив', '30'),
#     ('29', 'SSI-4330N0F', 'Наконечники на 1000 мкл, с фильтром  в пакете, 1000 шт/уп', '30'),
#     ('30', 'SSI-4237NSFS', 'Наконечники Vertex на 200 мкл, с фильтром, бесцветные, градуированные, стерильные, 96 шт/штатив', '100'),
#     ('31', 'SSI-4230N0F', 'Наконечники на 200 мкл, с фильтром  в пакете, 1000 шт/уп', '30'),
#     ('32', 'SSI-4137NSFS', 'Наконечники Vertex на 10 мкл Eppendorf/Biohit-совместимые, с фильтром, удлиненные, бесцветные, градуированные, стерильные, 96 шт/штатив', '3'),
#     ('33', 'SSI-4130N0F', 'Наконечники на 10 мкл, с фильтром  в пакете, Eppendorf/Biohit-совместимые, 1000 шт/уп', '30'),
#     ('34', 'SSI-1110-00', 'Пробирки 0,5 мл, градуированные,  1000 шт./уп.', '30'),
#     ('35', 'SSI-3225-00', 'Пробирки 0,2 мл тонкостенные с оптически-прозрачной плоской крышкой, 1000 шт/уп.', '30'),
#     ('36', 'SSI-3441-00', 'Планшеты тонкостенные для ПЦР низкопрофильные 0,1 мл, 96-луночные,Skirted, 10 шт/уп', '15'),
#     ('37', 'SSI-5500-29_', 'Штатив рабочий с крышкой для пробирок 0,2; 0,5 и 1,5-2,0 мл, цвет - ассорти и штатив для пробирок на 81 место', '4'),
#     ('38', 'SSI-5210-29', 'Штатив рабочий с крышкой для пробирок 0,2; 0,5 и 1,5-2,0 мл, цвет - ассорти', '10'),
# ]
# src = [
#     ("Клей для труб из ПВХ прозрачный 1 кг с кистью", "Tangit PVC-U", 12),
#     ("Очиститель 1 кг", "Tangit PVC-U/ABS", 12),
#     ("Нить для герметизации резьбовых соединений 100 м", "Tangit Uni-Lock", 4),
#     ("Муфта соединительная ПВХ FIP", "ø16", 4),
#     ("Муфта соединительная ПВХ FIP", "ø20", 5),
#     ("Муфта соединительная ПВХ FIP", "ø25", 15),
#     ("Муфта соединительная ПВХ FIP", "ø32", 40),
#     ("Муфта соединительная ПВХ FIP", "ø40", 5),
#     ("Муфта соединительная ПВХ FIP", "ø50", 6),
#     ("Муфта соединительная ПВХ FIP", "ø63", 3),
#     ("Муфта соединительная ПВХ FIP", "ø90", 1),
# ]
# src = [
#     ('8', '2.07.2.002', 'ПО Сервер МУ Контроль (v.2.0.1.lira)'),
#     ('9', '2.07.2.003', 'ПО Стандартная конфигурация интерфейса клиента "Главный монитор управления" '),
#     ('10', '2.07.2.004', 'ПО Сервер администрирования '),
#     ('11', '2.07.2.005', 'Руководство по эксплуатации к продукту'),
#     ('15', '2.07.2.009', 'ПО Дополнительная конфигурация интерфейса клиента "Главный монитор управления" -  Госпитальная платформа'),
#     ('19', '2.07.3.002', 'ПО Дополнительная конфигурация интерфейса клиента "Главный монитор управления" -  Обработка и конфигурация информации по инженерным сетям '),
#     ('21', '2.07.3.005', 'Модуль расширения-блок контроллера для медицинского оборудования.'),
#     ('22', '2.07.3.006', 'ПО Дополнительная конфигурация интерфейса клиента "Главный монитор управления" - контроль работы операционной лампы и стола'),
#     ('23', '2.07.3.007', 'ПО Дополнительная конфигурация интерфейса клиента "Главный монитор управления" - приемки и трансляции видеоданных, осуществление функций видеоинтеграции'),
#     ('24', '2.07.3.008', 'ПО Дополнительная конфигурация интерфейса клиента "Главный монитор управления" -  обработка  медицинских изображений '),
#     ('26', '2.07.3.021', 'ПО Дополнительная конфигурация интерфейса клиента "Главный монитор управления" - подключение вспомогательного монитора управления'),
#     ('27', '2.07.3.022', 'ПО Дополнительная конфигурация интерфейса клиента "Главный монитор управления" - подключение аудио устройств, медиатеки'),
#     ('29', '2.07.3.027', ' Система хранения и архивирования данных 20 ТБ'),
#     ('30', '2.07.3.028', 'ПО Дополнительная конфигурация интерфейса клиента "Главный монитор управления" - поиск и демонстрация архивных данных'),
#     ('31', '2.07.3.030', 'Модуль расширения- блок памяти  добавляет 40 ТБ'),
#     ('33', '2.07.3.033', 'ПО Дополнительная конфигурация интерфейса клиента "Главный монитор управления" -  организация видеоконференции между клиентами'),
#     ('35', '2.07.3.038', 'Модуль расширения- блок документооборота '),
#     ('36', '2.07.3.039', 'ПО Дополнительная конфигурация интерфейса клиента "Главный монитор управления" - подключение возможности сопряжения системы с МИС лечебного учреждения и работы с планом операций'),
# ]
# src = [
#     ('6', '2.07.2.002', 'ПО Сервер МУ Контроль (v.2.0.1.lira)'),
#     ('7', '2.07.2.003', 'ПО Стандартная конфигурация интерфейса клиента "Главный монитор управления" '),
#     ('8', '2.07.2.004', 'ПО Сервер администрирования '),
#     ('9', '2.07.2.005', 'Руководство по эксплуатации к продукту'),
#     ('12', '2.07.2.009', 'ПО Дополнительная конфигурация интерфейса клиента "Главный монитор управления" -  Госпитальная платформа'),
#     ('17', '2.07.3.002', 'ПО Дополнительная конфигурация интерфейса клиента "Главный монитор управления" -  Обработка и конфигурация информации по инженерным сетям '),
#     ('19', '2.07.3.005', 'Модуль расширения-блок контроллера для медицинского оборудования.'),
#     ('20', '2.07.3.006', 'ПО Дополнительная конфигурация интерфейса клиента "Главный монитор управления" - контроль работы операционной лампы и стола'),
#     ('21', '2.07.3.007', 'ПО Дополнительная конфигурация интерфейса клиента "Главный монитор управления" - приемки и трансляции видеоданных, осуществление функций видеоинтеграции'),
#     ('22', '2.07.3.008', 'ПО Дополнительная конфигурация интерфейса клиента "Главный монитор управления" -  обработка  медицинских изображений '),
#     ('26', '2.07.3.021', 'ПО Дополнительная конфигурация интерфейса клиента "Главный монитор управления" - подключение вспомогательного монитора управления'),
#     ('27', '2.07.3.022', 'ПО Дополнительная конфигурация интерфейса клиента "Главный монитор управления" - подключение аудио устройств, медиатеки'),
#     ('30', '2.07.3.027', ' Система хранения и архивирования данных 20 ТБ'),
#     ('31', '2.07.3.028', 'ПО Дополнительная конфигурация интерфейса клиента "Главный монитор управления" - поиск и демонстрация архивных данных'),
#     ('32', '2.07.3.029', 'Модуль расширения- блок памяти  добавляет 20 ТБ'),
#     ('35', '2.07.3.033', 'ПО Дополнительная конфигурация интерфейса клиента "Главный монитор управления" -  организация видеоконференции между клиентами'),
#     ('37', '2.07.3.038', 'Модуль расширения- блок документооборота '),
#     ('38', '2.07.3.039', 'ПО Дополнительная конфигурация интерфейса клиента "Главный монитор управления" - подключение возможности сопряжения системы с МИС лечебного учреждения и работы с планом операций'),
# ]

# levels = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
# level = [i for i in levels if i.Name == 'Этаж 01'][0]
# offset = 25
# sel = sel[0]
# y = sel.Location.Point.Y
# for pos in src:
#     symbols = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsElementType().ToElements()
#     symbol_name = '{} {}'.format(pos[0], ' '.join(pos[2].split()[:3]))
#     existing_symbol = [i for i in symbols if i.LookupParameter('Имя типа').AsString() == symbol_name]
#     if existing_symbol:
#         new_symbol = existing_symbol[0]
#     else:
#         new_symbol = doc.GetElement(sel.GetTypeId()).Duplicate(symbol_name)
#     # if 'new_symbol' in globals() and new_symbol.LookupParameter('Имя типа').AsString() == pos[0]:
#     #     pass
#     # else:
#     #     new_symbol = doc.GetElement(sel.GetTypeId()).Duplicate('{}'.format(pos[0]))
#     new_symbol.LookupParameter('Комментарии к типоразмеру').Set(pos[1])
#     new_symbol.LookupParameter('Стоимость').Set(float(pos[0]))
#     # new_symbol.LookupParameter('пиОсание').Set(pos[2])
#     # new_symbol.LookupParameter('Код по классификатору').Set(pos[1])
#     new_symbol.LookupParameter('Описание').Set(pos[2])
#     x = 0
#     el = doc.Create.NewFamilyInstance(XYZ(x * offset / k, -y * offset / k, 0), new_symbol, level, Structure.StructuralType.NonStructural)
#     # for i in range(int(pos[2])):
#     #     el = doc.Create.NewFamilyInstance(XYZ(x * offset / k, -y * offset / k, 0), new_symbol, level, Structure.StructuralType.NonStructural)
#     #     # el.LookupParameter('ХТ Размер фитинга ОВ').Set(pos[1])
#     #     x += 1
#     y += 1

# t.Commit()









# ids = []
# for el in sel:
#     name = WorksharingUtils.GetWorksharingTooltipInfo(doc, el.Id).LastChangedBy
#     ids.append(el.Id) if name != 'SG20' else None

# uidoc.Selection.SetElementIds(List[ElementId](ids))
# # for n in ids:
# #     print(n)


# # t.Commit()












# sum = 0
# for el in sel:
#   if not el.LookupParameter('Категория').AsValueString() == 'Помещения':
#       continue
#   number = el.LookupParameter('Номер').AsString()
#   el.LookupParameter('Комментарии').Set(number)
#   # area = el.LookupParameter('Площадь').AsDouble() / 45.088 * 4.189
#   # print('{:.1f}'.format(area))
#   # sum += area

# # print('---------')
# # print('{:.1f}'.format(sum))

# t.Commit()











# for el in sel:
#     sym = el.LookupParameter('Тип').AsValueString()
#     val = []
#     if '1х' in sym:
#         val.append(el.LookupParameter('1х').AsValueString())
#     if '2х' in sym:
#         val.append(el.LookupParameter('1х').AsValueString())
#         val.append(el.LookupParameter('2х').AsValueString())
#     if '3х' in sym:
#         val.append(el.LookupParameter('1х').AsValueString())
#         val.append(el.LookupParameter('2х').AsValueString())
#         val.append(el.LookupParameter('3х').AsValueString())
#     if '4х' in sym:
#         val.append(el.LookupParameter('1х').AsValueString())
#         val.append(el.LookupParameter('2х').AsValueString())
#         val.append(el.LookupParameter('3х').AsValueString())
#         val.append(el.LookupParameter('4х').AsValueString())
#     if '5х' in sym:
#         val.append(el.LookupParameter('1х').AsValueString())
#         val.append(el.LookupParameter('2х').AsValueString())
#         val.append(el.LookupParameter('3х').AsValueString())
#         val.append(el.LookupParameter('4х').AsValueString())
#         val.append(el.LookupParameter('5х').AsValueString())
#     if val:
#         # print('{}: {}: {}'.format(el.Id, sym, '; '.join(sorted(val))))
#         print('{}'.format('; '.join(sorted(val))))
#         el.LookupParameter('stuff').Set('{}'.format('; '.join(sorted(val))))
# t.Commit()








# els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType().ToElements()
# for el in els:
#     if el.LookupParameter('ADSK_Тепловая мощность'):
#         value = el.LookupParameter('ADSK_Тепловая мощность').AsDouble()
#         if value and value > 1:
#             el.LookupParameter('_').Set(value)
#         else:
#             value = el.LookupParameter('_').AsDouble()
#             if value and value > 1:
#                 el.LookupParameter('ADSK_Тепловая мощность').Set(value)
# t.Commit()











# from pyrevit.forms import ProgressBar

# elems = range(500)
# max_value = 100
# with ProgressBar() as pb:
#     # for counter in range(0, max_value):
#     #     pb.update_progress(counter, max_value)
#     for el in elems:


# t.Commit()











# src = [
#     ('Сетевой управляющий блок', '2.07.2.007'),
#     ('Органайзер для кабельного подключения', '2.07.2.001'),
#     ('Монитор управления главный 22" вертикальный', '2.07.1.004'),
#     ('Монитор управления главный 32"', '2.07.1.008'),
#     ('Динамик потолочный встраиваемый', '2.07.3.003'),
#     ('Монитор вспомогательный 49"', '2.07.1.018'),
#     ('Серверный телекоммуникационный шкаф 47U', '2.07.2.012'),
#     ('Видеокамера купольная, поворотная', '2.07.2.015'),
#     ('Микрофон беспроводной', '2.07.3.034'),
#     ('Комплект видеорозеток', '2.07.3.013'),
#     ('Станция рабочая 32" с клавиатурой', '2.07.3.045'),
#     ('Информационное табло 2×27"', '2.07.3.053'),
#     ('Система управления инженерными сетями', '2.07.3.001'),
#     ('Система хранения данных 20 ТБ', '2.07.3.027'),
#     ('Модуль расширения памяти 40 ТБ', '2.07.3.030'),
#     ('Система управления медицинским оборудованием', '2.07.3.005'),
#     ('Система мультимедиа на 16 каналов', '2.07.3.025'),
#     ('Система видеоменеджмента на 16 каналов', '2.07.3.011'),
#     ('Система видеоконференции', '2.07.3.032'),
#     ('Система документооборота', '2.07.3.038'),
# ]

# n = 1
# y = sel.Location.Point.Y
# for pos in src:
#     new_symbol = doc.GetElement(sel.GetTypeId()).Duplicate('{} {}'.format(pos[1], pos[0]))
#     new_symbol.LookupParameter('ADSK_Наименование').Set(pos[0])
#     new_symbol.LookupParameter('ADSK_Позиция').Set(str(n))
#     new_symbol.LookupParameter('ADSK_Код изделия').Set(pos[1])
#     el = doc.Create.NewFamilyInstance(XYZ(0, -y * 15 / k, 0), new_symbol, doc.ActiveView)
#     y += 1
#     n += 1

# t.Commit()












# t.SetName('Test: Заполнение названий и площадей пространств')
# src = {

#     6.103: ('Пример №1', 0),
#     6.104: ('Пример №2', 0),
#     6.105: ('Пример №3', 0),
# }


# src = {
#     101: ('Кабинет мастеров', 0),
#     102: ('С/у (м)', 0),
#     103: ('С/у (ж)', 0),
#     104: ('Кабинет', 0),
#     105: ('Раздевалка', 0),
#     106: ('Тамбур', 0),
#     107: ('Душе бая', 0),
#     108: ('с/у', 0),
#     109: ('Прачечная', 0),
#     110: ('Серверная', 0),
#     111: ('Выставочный зал', 0),
#     112: ('Вестибюль', 0),
#     113: ('КПП', 0),
#     114: ('Лестничная клетка', 0),
#     115: ('Проектный отдел', 0),
#     116: ('Конструкторский отдел', 0),
#     117: ('Комната приёма пищи', 0),
#     118: ('Электрощитовая', 0),
#     119: ('Комната охраны труда (мед. пункт)', 0),
#     120: ('ИТП + водомерный узел', 0),
#     121: ('Коридор', 0),
#     122: ('Подсобное помещение', 0),
#     123: ('КУИ', 0),
#     124: ('Раздевалка', 0),
#     125: ('Тамбур', 0),
#     126: ('Душевая', 0),
#     127: ('С/у', 0),
#     128: ('Кладовая', 0),
#     129: ('Подсобное помещение', 0),
#     130: ('С/у (м)', 0),
#     131: ('С/у (ж)', 0),
#     132: ('Подсобное помещение', 0),
#     133: ('Кафетерий', 0),
#     134: ('Офис', 0),
#     135: ('Кабинет', 0),
#     136: ('Лестничная клетка', 0),
#     137: ('Офис', 0),
#     138: ('Офис', 0),
#     139: ('Приемная', 0),
#     140: ('Коридор', 0),
#     141: ('Зарядная', 0),
#     142: ('Подсобное помещение', 0),
#     143: ('Склад', 0),
#     144: ('Склад', 0),
#     145: ('Зона открытого склада', 0),
#     201: ('Кабинет', 0),
#     202: ('Кухня', 0),
#     203: ('С/у (м)', 0),
#     204: ('С/у (ж)', 0),
#     205: ('Комната приёма пищи', 0),
#     206: ('Комната отЗыха', 0),
#     207: ('Конференц-зал', 0),
#     208: ('КориЗор', 0),
#     209: ('Офис', 0),
#     210: ('Лестничная клетка', 0),
#     211: ('Кабинет исполнительного Директора', 0),
#     212: ('Приёмная', 0),
#     213: ('Кабинет генерального Директора', 0),
#     214: ('Кабинет', 0),
#     215: ('Кабинет', 0),
#     216: ('Кабинет', 0),
#     217: ('Конференц-зал', 0),
#     218: ('Серверная', 0),
#     219: ('ПоЗсобное помещение', 0),
#     220: ('С/у(м)', 0),
#     221: ('С/у (ж)', 0),
#     222: ('КУИ', 0),
#     223: ('Служебное помещение', 0),
#     224: ('Кабинет', 0),
#     225: ('Офис', 0),
#     226: ('Кабинет', 0),
#     227: ('Лестничная клетка', 0),
#     228: ('Приёмная', 0),
#     229: ('Кабинет', 0),
#     230: ('Офис', 0),
#     231: ('КориЗор', 0),
#     232: ('Холл', 0),
# }
# spaces = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElements()

# for el in spaces:
#     num = float(el.LookupParameter('Номер').AsString())
#     if num in src:
#         el.LookupParameter('Имя').Set(src[num][0])
#         # el.LookupParameter('Площадь_').Set(str(src[num][1]))
#         # el.LookupParameter('Площадь_').Set(src[num][1])

# t.Commit()






















# t = Transaction(doc, 'Test')
# t.Start()

# sel = [doc.GetElement(elid) for elid in uidoc.Selection.GetElementIds()]

# familys = list(set([doc.GetElement(el.GetTypeId()).Family for el in sel]))

# for i in familys:
#   print(i.Name)
#   doc.Delete(i.Id)


# t.Commit()



























# elects = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().ToElements()
# anns = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericAnnotation).WhereElementIsNotElementType().ToElements()
# anns = [i for i in anns if i.LookupParameter('Семейство').AsValueString() == 'Элемент схемы']
# dict_ann = {}
# for el in anns:
#     symbol = doc.GetElement(el.GetTypeId())
#     position = symbol.LookupParameter('ADSK_Позиция').AsString()
#     if position not in dict_ann:
#         dict_ann[position] = []
#     dict_ann[position].append(el)

# dict_elect = {}
# for el in elects:
#     symbol = doc.GetElement(el.GetTypeId())
#     position = symbol.LookupParameter('Стоимость').AsDouble()
#     if position > 100:
#         continue
#     position = '{0:n}'.format(position).replace(',', '.')
#     if position not in dict_elect:
#         dict_elect[position] = []
#     dict_elect[position].append(el)
# # print(dict_elect)

# data = []
# limit = 274  # Вероятная проблема при пустом выводе
# positions = set(dict_ann.keys() + dict_elect.keys())
# for position in natural_sorted(positions):
#     row = []
#     row.append(position)
#     els = dict_ann.get(position, [])
#     row.append('; '.join(set([doc.GetElement(i.GetTypeId()).LookupParameter('ADSK_Наименование').AsString() for i in els])))
#     ann_len = len(els)
#     row.append(ann_len)
#     els = dict_elect.get(position, [])
#     descriptions = []
#     for i in els:
#         description = doc.GetElement(i.GetTypeId()).LookupParameter('Описание').AsString()
#         if '\r' in description:
#             description = description.split('\r')[0] + ' …'
#         if len(description) > limit:
#             description = description[:limit] + '…'  # Вероятная проблема при пустом выводе
#         descriptions.append(description)
#     row.append('; '.join(set(descriptions)))
#     row.append(len(els))
#     # diff = ann_len - len(els)
#     row.append(ann_len - len(els) or '')

#     data.append(row)

# output = script.get_output()

# left = [1, 3]
# for row in data:
#     formatted_row = []
#     for index, item in enumerate(row):
#         formatted_row.append(str(item)[:50].ljust(50, ' ') if index in left else str(item).center(10, ' '))
#     out = ';'.join(formatted_row)
#     print(out)
#     # output.print_html('<h1>' + out + '</h1>')
#     # print(':::'.join([str(i)[:50] for i in row]))


# total2 = sum([i[2] for i in data])
# total4 = sum([i[4] for i in data])
# total5 = len([i[5] for i in data if i[5]])
# data.append(['Итого: ' + str(len(data)),
#              '---',
#              total2,
#              '---',
#              total4,
#              total5,
#              ])

# columns = ['Позиция', 'Наименование по схеме', 'Количество', 'Наименование по спеке', 'Количество', 'Разница']

# output.print_table(table_data=data, columns=columns)

# t.Commit()

# # output.print_html('<h1>' + out + '</h1>')




























# src = [(1824740, ''), (1824920, ''), (1824922, ''), (1824924, ''), (1824926, ''), (1824928, ''), (1824930, ''), (1824932, ''), (1824934, ''), (1826765, ''), (1836540, ''), (1836542, ''), (1836544, ''), (1836546, ''), (1836548, ''), (1836550, ''), (1836552, ''), (1836554, ''), (2475322, 'utp/utp'), (2604386, 'sdi'), (2700046, ''), (2700048, ''), (2700050, ''), (2700052, ''), (2700054, ''), (2806834, 'hdmi'), (2807028, 'vga'), (2948151, 'utp/utp'), (2948153, 'hdmi'), (2948155, 'utp'), (2948157, 'hdmi((2))/usb/pwr/pwr'), (2948159, 'utp'), (2948161, 'utp/utp'), (2953615, 'hdmi((2))/usb'), (2957709, 'utp'), (2957711, 'pwr'), (2957713, 'pwr'), (2957715, 'ftp'), (2957717, 'pwr'), (2957719, 'pwr'), (2957721, 'audio'), (2963054, 'pwr'), (2963056, 'pwr'), (2966584, 'utp/sdi'), (2966586, 'utp/sdi'), (2981279, ''), (2981280, ''), (2981281, ''), (2981282, ''), (2981283, ''), (2981284, ''), (2981285, ''), (2981286, ''), (2981287, ''), (2981365, 'utp((50))'), (3005435, 'pwr/pwr'), (3005863, '((50))utp/ftp/pwr'), (3006081, 'utp'), (3027840, 'hdmi'), (3027842, 'usb'), (3027844, 'utp'), (3027846, 'sdi'), (3027848, 'sdi'), (3027850, 'hdmi'), (3027852, 'sdi'), (3031538, 'hdmi/vga/sdi/sdi/4alarm'), (3031939, 'sdi'), (3033219, 'sdi'), (3034576, 'hdmi'), (3034725, 'vga'), (3090131, 'utp'), (3090201, 'audio'), (3090203, 'audio'), (3097521, 'Внимание! Если написать сюда 64 раза sdi, то Ревиту станет плохо. Надо что-то думать'), (3097523, 'sdi/sdi/sdi/sdi/sdi/sdi/sdi/sdi/sdi/sdi/sdi/sdi/sdi/sdi/sdi/sdi/sdi/sdi/sdi/sdi/sdi/sdi/sdi/sdi/sdi/sdi/sdi/sdi/sdi/sdi/sdi/sdi'), (3097525, 'sdi/sdi/sdi/sdi/sdi/sdi/sdi/sdi/sdi/sdi/sdi/sdi/sdi/sdi/sdi/sdi'), (3106722, 'utp/audio/hdmi'), (3106724, 'SDI/SDI/SDI/SDI/utp/audio'), (3106726, 'audio'), (3106728, 'audio'), (3106730, 'utp'), (3106732, 'utp'), (3106734, 'utp'), (3107242, 'utp'), (3107386, ''), (3128209, ''), (3137253, 'audio'), (3137457, 'sdi'), (3153566, ''), (3153789, ''), (3153804, ''), (3153829, ''), (3169041, ''), (3186155, ''), (3186182, ''), (3232020, 'utp'), (3232179, 'utp'), (3233011, ''), (3253312, 'utp'), (3253864, 'hdmi/hdmi/vga/sdi/4alarm/utp')]
# # els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsElementType().ToElements()
# for i in src:
#   doc.GetElement(ElementId(i[0])).LookupParameter('КабельСигнальный').Set(i[1])
#   # print('{}\t{}'.format(symbol.Id, symbol.LookupParameter('КабельСигнальный').AsString()))
# t.Commit()




# schedules = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Schedules).WhereElementIsNotElementType().ToElements()

# schedules = list(filter(lambda x: 'Эл ' in x.Name, schedules))

# for sch in schedules:
#     val = float(sch.Name.split('Эл ')[-1].replace(',', '.'))
#     for fieldindex, sch_filter in enumerate(sch.Definition.GetFilters()):
#         if sch_filter.IsDoubleValue:
#             if sch_filter.GetDoubleValue() == 1.01:
#                 sch_filter.SetValue(val)
#             elif sch_filter.GetDoubleValue() == 1.011:
#                 sch_filter.SetValue(val + 0.001)
#         sch.Definition.SetFilter(fieldindex, sch_filter)

# t.Commit()




# grids = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Grids).WhereElementIsNotElementType().ToElements()
# for el in grids:
#     el.SetVerticalExtents(-50, 150)


# t.Commit()


# def natural_sorted(list, key=lambda s: s):
#     """
#     Sort the list into natural alphanumeric order.
#     """
#     def get_alphanum_key_func(key):
#         convert = lambda text: int(text) if text.isdigit() else text  # noqa
#         return lambda s: [convert(c) for c in re.split('([0-9]+)', key(s))]
#     sort_key = get_alphanum_key_func(key)
#     return sorted(list, key=sort_key)


# views = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views).WhereElementIsNotElementType().ToElements()
# views = [i for i in views if i.Origin]
# views = [i for i in views if i.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString() == '3D вид']
# views = natural_sorted(views, key=lambda x: x.Name)
# views = sorted(views, key=lambda x: x.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString())
# # data = []
# for view in views:
#     # view_filters = [doc.GetElement(ElementId(i)) for i in [1187542, 1189543, 1189548, ]]
#     for elid, pat in [(1187542, 'ADSK_Накрест косая_2мм'), (1189543, 'General_Honeycomb'), (1189548, 'STARS'), ]:
#         filt = doc.GetElement(ElementId(elid))
#         cfg = view.GetFilterOverrides(filt.Id)
#         # cfg.SetProjectionLineWeight(              int(current_pro_line_weight)           ) if      int(current_pro_line_weight)     > 0 else None  # noqa
#         # cfg.SetProjectionLineColor(               col(current_pro_line_color)            ) if      col(current_pro_line_color)          else None  # noqa
#         # cfg.SetProjectionLinePatternId(          line(current_pro_line_pattern_id)       ) if     line(current_pro_line_pattern_id)     else None  # noqa
#         # cfg.SetProjectionFillPatternVisible( bool(int(current_pro_fill_pattern_visible)) )                                                         # noqa
#         cfg.SetProjectionFillPatternId(          fill(pat)                               )
#         cfg.SetProjectionFillColor(               col('192 192 192')                     )
#         # cfg.SetSurfaceTransparency(               int(current_transparency)              )                                                         # noqa
#         # cfg.SetCutLineWeight(                     int(current_cut_line_weight)           ) if      int(current_cut_line_weight)     > 0 else None  # noqa
#         # cfg.SetCutLineColor(                      col(current_cut_line_color)            ) if      col(current_cut_line_color)          else None  # noqa
#         # cfg.SetCutLinePatternId(                 line(current_cut_line_pattern_id)       ) if     line(current_cut_line_pattern_id)     else None  # noqa
#         # cfg.SetCutFillPatternVisible(        bool(int(current_cut_fill_pattern_visible)) )                                                         # noqa
#         # cfg.SetCutFillPatternId(                 fill(current_cut_fill_pattern_id)       ) if     fill(current_cut_fill_pattern_id)     else None  # noqa
#         # cfg.SetCutFillColor(                      col(current_cut_fill_color)            ) if      col(current_cut_fill_color)          else None  # noqa
#         # cfg.SetHalftone(                     bool(int(current_halftone))                 )                                                         # noqa
#         # view.SetFilterOverrides(ElementId(int(current_filter_id)), cfg)


# els = FilteredElementCollector(doc).WhereElementIsElementType().ToElements()
# for el in els:
#     if el.LookupParameter('Изготовитель'):
#         if el.LookupParameter('Изготовитель').HasValue:
#             old = el.LookupParameter('Изготовитель').AsString() == 'Лиссант'
#             if old:
#                 el.LookupParameter('Изготовитель').Set('Россия')
#                 print('{} {}'.format(old, el.LookupParameter('Имя типа').AsString()))


# sel = [doc.GetElement(elid) for elid in uidoc.Selection.GetElementIds()]
# ducts = [i for i in sel if i.LookupParameter('Категория').AsValueString() == 'Воздуховоды']

# done = []
# for el in ducts:
#   sys = el.MEPSystem
#   if sys:
#       if sys.Id.IntegerValue not in done:
#           sys.Name += ' транзит 7'
#           done.append(sys.Id.IntegerValue)

# t.Commit()

# cir0 = [i for i in sel if i.Name == 'Линии детализации'][0]
# cir1 = [i for i in sel if i.Name == 'Линии детализации'][1]
# # view = [i for i in sel if i.Name != 'Линии детализации'][0]

# t = Transaction(doc, 'Test')
# t.Start()

# # for i in dir(view.Location):
# #     print(i)
# # cir.Location.Move(view.GetBoxCenter())
# print(doc.ActiveView.Outline.Min)
# print(doc.ActiveView.Outline.Max)
# # cir0.Location.Move(doc.ActiveView.Outline.Min - cir0.Location.Curve.Center)
# # cir0.Location.Move(doc.ActiveView.Outline.Max - cir1.Location.Curve.Center)
# # cir0.Location.Move(view.GetBoxOutline().MinimumPoint - cir0.Location.Curve.Center)
# # cir1.Location.Move(view.GetBoxOutline().MaximumPoint - cir1.Location.Curve.Center)
# # view.GetBoxOutline().MinimumPoint
# # view.GetBoxOutline().MaximumPoint


# t.Commit()

# # filters = [doc.GetElement(i) for i in doc.ActiveView.GetFilters()]
# # for i in filters:
# #   print(i.Name)
# #   print(doc.ActiveView.GetFilterVisibility(i.Id))
# #   print(doc.ActiveView.GetFilterOverrides(i.Id).Halftone)
# #   color = doc.ActiveView.GetFilterOverrides(i.Id).ProjectionLineColor
# #   if color.IsValid:
# #       rgb = (color.Red, color.Green, color.Blue)
# #       output.print_html('{} <font color="rgb({},{},{})">██████████</font>'.format(rgb, rgb[0], rgb[1], rgb[2]))
# #   else:
# #       output.print_html('Uninitiaized')
# #   print()

# all_filters = FilteredElementCollector(doc).OfClass(ParameterFilterElement).ToElements()
# # for i in all_filters:
# #   print(i.Name)

# t = Transaction(doc, 'Test')
# t.Start()

# names = ['В1', 'В2', 'В3', 'В12', 'В20', 'В172', 'ДУ5', 'К1', 'К2', 'К3', 'К12', 'К20', 'П30', 'ПД16', ]

# cats = List[ElementId]([
#     ElementId(BuiltInCategory.OST_DuctAccessory),
#     ElementId(BuiltInCategory.OST_DuctCurves),
#     ElementId(BuiltInCategory.OST_DuctFitting),
#     ElementId(BuiltInCategory.OST_DuctInsulations),
#     ElementId(BuiltInCategory.OST_DuctTerminal),
#     ElementId(BuiltInCategory.OST_FlexDuctCurves),
# ])


# for name in names:
#     pfilter = parameterFilterElement = ParameterFilterElement.Create(
#         doc,
#         name + ' не',
#         cats,
#         ElementParameterFilter(ParameterFilterRuleFactory.CreateNotEqualsRule(ElementId(BuiltInParameter.RBS_SYSTEM_NAME_PARAM), name, True))
#     )
#     # doc.ActiveView.Duplicate(name).SetFilterVisibility([i.Id for i in all_filters if i.Name == name + ' не'][0], False)
#     print(name)
#     new_view = doc.GetElement(doc.ActiveView.Duplicate(ViewDuplicateOption.Duplicate))
#     new_view.Name = name
#     new_view.SetFilterVisibility(pfilter.Id, False)

# # doc.ActiveView.SetFilterVisibility([i.Id for i in all_filters if i.Name == 'К1 не'][0], False)

# t.Commit()

# els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).WhereElementIsNotElementType().ToElements()
# sel = [doc.GetElement(elid) for elid in uidoc.Selection.GetElementIds()]
# sel = [i.LookupParameter('Этап').AsDouble() for i in sel if i.LookupParameter('Категория').AsValueString() == 'Каркас несущий']
# el_ids = List[ElementId]([i.Id for i in els if i.LookupParameter('Этап').AsDouble() in sel])

# uidoc.Selection.SetElementIds(el_ids)



# for el in els:
#   s = '{}°°°{}°°°{}°°°{}°°°{}°°°{}°°°{}°°°{}°°°{}°°°{}°°°'.format(
#       el.LookupParameter('Имя типа').AsString(),
#       el.LookupParameter('Описание').AsString(),
#       el.LookupParameter('Комментарии к типоразмеру').AsString(),
#       el.LookupParameter('URL').AsString(),
#       el.LookupParameter('Изготовитель').AsString(),
#       el.LookupParameter('Группа модели').AsString(),
#       el.LookupParameter('Ключевая пометка').AsString(),
#       el.LookupParameter('КабельСигнальный').AsString(),
#       el.LookupParameter('КабельСиловой').AsString(),
#       el.LookupParameter('Изображение типоразмера').AsValueString(),
#   )
#   print(s)
#        0  1           2                       3                                   4                   5               6                           7                           8                   9               10                      11
# src = """ОК   Имя типа    Наименование (Описание) Комментарии к типоразмеру (Артикул) URL (УГО Описание)  Изготовитель    Группа модели (примечание)  Ед. изм. (Ключевая пометка) КабельСигнальный    КабельСиловой   Изображение типоразмера -
# 0 MiniPC  Не включать                     UTP 4x2x0,52 cat 5e         -
# 0 PACS1   Не включать Артикул                 UTP 2x2x0,52 cat 5e         -
# 2 БААВС HDMI, VGA, SDI, CVBS  Блок аппаратной адаптации видеосигнала (БААВС) HDMI, VGA, SDI, CVBS Артикул Блок аппаратной адаптации видеосигнала (БААВС)  ООО "Медицинские системы визуализации"  Блок аппаратной адаптации видеосигнала (БААВС) HDMI, VGA, SDI, CVBS шт. RG-6U (75 Ом)       42 БААВС.png    -
# 3 Беспроводной микрофон с клипсой Микрофон беспроводной с поясным модулем и клипсой   1.07.4.004  Без УГО ООО "Медицинские системы визуализации"  Беспроводной микрофон необходим для передачи и записи звуков в операционной. Модель микрофона предназначена для индивидуального использования - крепится на одежде основного говорящего. В комплект поставки включены беспроводной приёмник и зарядное устройство для него. шт.             -
# 1 Блок контроля освещения (DALI)  Не включать     Блок контроля освещения (DALI) с указанием числа светильников               ШВВП 2х0,75 По кратчайшему пути 30.png  -
# 1 Блок контроля освещения (DALI) значок выкл  Не включать     Блок контроля освещения (DALI) с указанием числа светильников               ШВВП 2х0,75 По кратчайшему пути 30.png  -
# 2 Блок розеток HDMI, VGA, SDI, CVBS   Блок из четырех розеток HDMI, VGA, SDI, CVBS    Артикул Без УГО ООО "Медицинские системы визуализации"      шт. HDMI/HDMI/RG-6U (75 Ом)/RG-6U (75 Ом)/4×Alarm/HDMI/VGA/2×Alarm  По кратчайшему пути     -
# 1 Блок сетевой управляющий ВЕГА   Блок сетевой управляющий (БСУ ВЕГА) с предустановленным базовым программным обеспечением    1.07.1.004  Блок сетевой управляющий ВЕГА   ООО "Медицинские системы визуализации"  Блок с предустановленным программным обеспечением предназначен для управления всеми функциями системы. В комплект поставки включены датчики температуры, влажности и реле давления, которые позволяют следить за климатическими параметрами в операционной и управлять вентиляционной системой. Герметичное исполнение блока обеспечивает надежность и безопасность работы. Компактный размер блока дает возможность разместить его за обшивкой стен или потолка и сэкономить место в серверной комнате.    шт. UTP 4x2x0,52 cat 5e     1.png   -
# 1 Видеоисточник HDMI  Не включать     Вывод кабеля для подключения видеоисточника с указанием интерфейса              HDMI    По кратчайшему пути 43 Вывод кабеля.png -
# 1 Видеоисточник VGA   Не включать     Вывод кабеля для подключения видеоисточника с указанием интерфейса              VGA По кратчайшему пути 43 Вывод кабеля.png -
# 2 Видеоисточник SDI   Не включать     Вывод кабеля с указанием типа               RG-6U (75 Ом)           -
# 2 Датчик перепада давления    Датчик загрязнения HEPA-фильтра системы вентиляции (входит в комплект БСУ, не является отдельной позицией)  Артикул Датчик загрязнения HEPA-фильтра системы вентиляции  ООО "Медицинские системы визуализации"      шт. ШВВП 2х0,75     12.png  -
# 2 Датчик температуры и влажности  Датчик температуры и влажности воздуха (входит в комплект БСУ, не является отдельной позицией)  Артикул Датчик температуры и влажности воздуха  ООО "Медицинские системы визуализации"      шт. FTP 4x2x0,52 cat 5e     11.png  -
# 1 Динамик потолочный  Динамик потолочный встраиваемый 1.07.4.005  Динамик потолочный  ООО "Медицинские системы визуализации"  Динамик необходим для трансляции звуковых сигналов сигнализации и состояния оборудования. Динамик предназначен для встраивания в потолок. Диаметр: 120 мм   шт. Акустический кабель Audiocore Primary Wire M ACS0102        17.png  -
# 1 Конвертер HDMI to SDI   Конвертер видеосигналов HDMI to SDI 1.07.5.006  Конвертер видеосигнала HDMI to SDI  ООО "Медицинские системы визуализации"  Конвертер предназначен для аппаратного преобразования и масштабирования потокового видеосигнала с аппаратных источников различного формата в SDI. Конвертер необходим для интеграции в систему оборудования с видеовыходами различного формата. шт. RG-6U (75 Ом)       9.png   -
# 1 Конвертер UTP to USB    Не включать     Конвертер USB to RJ45 to USB                USB 2.0     6.png   -
# 2 Конвертер SDI to HDMI   Конвертер видеосигналов SDI to HDMI Артикул Конвертер SDI to HDMI   ООО "Медицинские системы визуализации"      шт. ((2))HDMI       25.png  -
# 1 Конвертер USB to UTP    Удлинитель USB интерфейса (в комплекте приемник и передатчик)   1.07.1.044  Конвертер USB to RJ45 to USB    ООО "Медицинские системы визуализации"  Удлинитель USB интерфейса необходим для соединения БСУ и монитора управления    шт. UTP 4x2x0,52 cat 5e     6.png   -
# 1 Конвертер  VGA to HDMI  Конвертер видеосигналов VGA/HDMI    1.07.5.005  Конвертер видеосигнала VGA to HDMI  ООО "Медицинские системы визуализации"  Конвертер предназначен для аппаратного преобразования и масштабирования потокового видеосигнала с аппаратных источников различного формата в HDMI. Конвертер необходим для интеграции в систему видеоменеджмента оборудования, имеющего видеовыход формата HDMI шт. HDMI    По кратчайшему пути 10.png  -
# 1 Контрольно-отключающая коробка МГС (КОК)    Не включать     Контрольно-отключающая коробка медицинского газоснабжения               UTP 4x2x0,52 cat 5e/UTP 4x2x0,52 cat 5e     39 КОК.png  -
# 3 Кронштейн на стену  Кронштейн для крепления БСУ ВЕГА на стену   1.07.1.045  Без УГО ООО "Медицинские системы визуализации"  Комплект крепежных элементов для размещения блока управления на стене   шт.             -
# 3 Модуль расширения памяти - 20 Тб    Модуль расширения памяти    1.07.7.001  Без УГО ООО "Медицинские системы визуализации"  Модуль необходим для увеличения объема хранения информации до 20 ТБ.    шт.             -
# 3 Модуль расширения. Блок видео. 4 вход/выход Модуль расширения - блок видео  1.07.5.001  Без УГО ООО "Медицинские системы визуализации"  Блок-видео предназначен для расширения возможностей системы функцией приемки и обработки видеопотоков. С помощью этого блока организована система сбора видеоизображений с источников видеосигнала. шт.             -
# 3 Модуль расширения. Блок видеоконференции и транскодирования Модуль расширения - блок видеоконференции и транскодирования    1.07.8.004  Без УГО ООО "Медицинские системы визуализации"  Блок-видео предназначен для расширения возможностей системы функцией приемки и обработки видеопотоков. С помощью этого блока происходит коммутация и компиляция, декодирование и передача видеосигналов, организация телекоммуникационного интерактивного взаимодействия удаленных абонентов со средствами трансляции аудио- и видеопотоков внутри операционного помещения  шт.             -
# 3 Модуль расширения. Блок документооборота (DICOM/HL7)    Модуль расширения - блок документооборота (DICOM/HL7)   1.07.8.001  Без УГО ООО "Медицинские системы визуализации"  Блок необходим для интеграции системы с документооборотом медицинского учреждения в формате DICOM и HL7. Система позволяет документировать процессы лечения, формировать и отправлять отчеты в систему МИС. Организовывает единую сеть с системой МИС, использует специальные протоколы для обмена информацией. шт.             -
# 3 Модуль расширения. Блок контроллера для инженерных сетей    Модуль расширения - блок контроллера для инженерных сетей   1.07.2.001  Без УГО ООО "Медицинские системы визуализации"  Блок коммутационный предназначен для контроля и управления параметрами инженерных систем, используется для сбора, хранения и обработки данных, поступающих с инженерных госпитальных систем и является необходимым элементом системы. Коммутационный блок оснащен программируемым контроллером для управления инженерными сетями и имеет 32 дискретных и аналоговых выхода. шт.             -
# 3 Модуль расширения. Блок медиа   Модуль расширения - блок медиа  1.07.4.001  Без УГО ООО "Медицинские системы визуализации"  Модуль необходим для обеспечения возможности приемки и обработки аудиопотоков, а также их трансляции в операционной. Благодаря этому элементу обеспечивается возможность воспроизводить внутри операционных аудиозаписи со встроенных носителей аудио информации, ресурсов хранения во внутрибольничной сети; вести запись переговоров между врачами.   шт.             -
# 1 Монитор коридорный 27"  Монитор управления вспомогательный 27" сенсорный, медицинское исполнение    1.07.1.072  Монитор управления вспомогательный сенсорный с вебкамерой   ООО "Медицинские системы визуализации"  Монитор управления устанавливается за пределами операционной (в соседнем помещении, например, в коридоре) для организации и управления видеоконференцией с операционной. Сенсорный монитор управления укомплектован WEB камерой, динамиками и микрофоном. Размер монитора по диагонали 27"  шт. UTP 4x2x0,52 cat 5e     21.png  -
# 1 Монитор обзорный 49"    Монитор медицинский специальный, диагональ 49"  1.07.1.082  Монитор медицинский с указанием размера ООО "Медицинские системы визуализации"  Обзорный медицинский монитор, устанавливаемый в операционной для трансляции изображения с видеоисточников. Размер монитора по диагонали: 49". Монитор оснащен защитным прозрачным сенсорным экраном в медицинском герметичном исполнении.   шт. UTP 4x2x0,52 cat 5e/HDMI        5.png   -
# 2 Монитор специальный для инфо-табло 2×24"    Монитор медицинский специальный сдвоенный (инфопанель) 2×24"    Артикул Инфопанель 2×24"    ООО "Медицинские системы визуализации"      шт. UTP 4x2x0,52 cat 5e/UTP 4x2x0,52 cat 5e     43 Инфопанель.png   -
# 1 Монитор управления 27" главный  Монитор управления главный, сенсорный, медицинское герметичное исполнение, диагональ 27"    1.07.1.005  Монитор управления сенсорный с указанием размера    ООО "Медицинские системы визуализации"  Главный монитор управления, располагаемый в операционной для управления всеми функциями системы является неотъемлемой частью системы MVS VEGA. На мониторе отображаются элементы управления и информационные виджеты. Медицинский сенсорный монитор в герметичном исполнении удобен для работы в условиях операционной и не боится обработки дезинфицирующими растворами. Размер монитора по диагонали: 27". Монитор оснащен защитным прозрачным сенсорным экраном в медицинском герметичном исполнении.    шт. HDMI((2))/USB 2.0/ШВВП 2х0,75/ШВВП 2х0,75       20.png  -
# 1 Монитор управления 32" главный  Монитор управления главный, сенсорный, медицинское герметичное исполнение, диагональ 32"    1.07.1.006  Монитор управления сенсорный с указанием размера    ООО "Медицинские системы визуализации"  Главный монитор управления, располагаемый в операционной для управления всеми функциями системы является неотъемлемой частью системы MVS VEGA. На мониторе отображаются элементы управления и информационные виджеты. Медицинский сенсорный монитор в герметичном исполнении удобен для работы в условиях операционной и не боится обработки дезинфицирующими растворами. Размер монитора по диагонали 32". Монитор оснащен защитным прозрачным сенсорным экраном размером (ШхВ) 992х792 мм, медицинское герметичное исполнение шт. HDMI/USB 2.0/ШВВП 2х0,75/ШВВП 2х0,75        20.png  -
# 1 Обзорная видеокамера    Видеокамера купольная поворотная    1.07.1.033  Видеокамера купольная, поворотная   ООО "Медицинские системы визуализации"  Сетевая PTZ-камера AXIS V5915. Видеокамера предназначена для трансляции потокового видео из операционной. Модель позволяет получать видео высокой четкости с разрешением HDTV 1080p. Поддерживается утилита AXIS Streaming Assistant. Защита с помощью паролей, фильтрация IP-адресов, HTTPS кодирование, управление доступом в сеть IEEE 802.1X, дайджест-проверка подлинности, журнал доступа пользователей. Открытый интерфейс API для интеграции программного обеспечения, включая VAPIX® и AXIS Camera Application Platform компании Axis Communications   шт. UTP 4x2x0,52 cat 5e/RG-6U (75 Ом)       15.png  -
# 3 Органайзер  Органайзер для кабельного подключения для БСУ ВЕГА  1.07.1.047  Без УГО ООО "Медицинские системы визуализации"  Органайзер предназначен для компактной укладки кабелей между установленным активным оборудованием и коммутационными панелями. Позволяет избегать спутывания и перекручивания проводов. Позволяет производить укладку кабелей с соблюдением требований по допустимым радиусам изгиба шт.             -
# 1 Рабочая станция 27" с клавиатурой   Станция рабочая, медицинское исполнение, клавиатура, манипуляционная панель, за защитным стеклом в едином исполнении    1.07.8.003  Рабочая станция с клавиатурой   ООО "Медицинские системы визуализации"  Рабочее место медсестры предназначено для организации видеоконференции. Рабочее место состоит из системного блока, сенсорного монитора и клавиатуры в защитном корпусе. Системный блок: 8 GB DDR4 RAM, 500 GB HDD. Размер монитора по диагонали 27". Сенсорный монитор располагается на стене операционной в удобном для работы месте и закрыт защитным фронтальным прочным стеклом. Размер фронтальной панели: 900 мм х 993 мм. Защитный корпус клавиатуры предназначен для обработки дезинфицирующими растворами  шт. UTP 4x2x0,52 cat 5e     22.png  -
# 1 Розетка SDI Блок аппаратной адаптации видеосигнала BNC (SDI)    1.07.8.011  Розетка для подключения видеоисточника с указанием интерфейса   ООО "Медицинские системы визуализации"  Розетка BNC (SDI)   шт. RG-6U (75 Ом)       7.png   -
# 1 Розетка HDMI    Блок аппаратной адаптации видеосигнала HDMI 1.07.8.012  Розетка для подключения видеоисточника с указанием интерфейса   ООО "Медицинские системы визуализации"  Розетка HDMI    шт. HDMI        7.png   -
# 1 Розетка VGA Блок аппаратной адаптации видеосигнала VGA  1.07.8.013  Розетка для подключения видеоисточника с указанием интерфейса   ООО "Медицинские системы визуализации"  Розетка VGA шт. VGA     7.png   -
# 1 Сервер  Сервер госпитальный 1.07.10.002 Без УГО ООО "Медицинские системы визуализации"  Необходим для обеспечения возможности хранения получаемых медицинских данных.   шт. UTP 4x2x0,52 cat 5e         -
# 1 Табло "Не входить"  Не включать     Табло "Не входить"              ШВВП 2х0,75     40 Табло Не входить.png -
# 1 Табло "Рентген" Не включать     Табло "Рентген"             ШВВП 2х0,75     41 Табло Рентген.png    -
# 1 Блок управления приводом двери  Не включать     Блок управления приводом двери              UTP 4x2x0,52 cat 5e     14.png  -
# 1 Блок управления приводом жалюзи Не включать     Блок управления приводом жалюзи             ШВВП 2х0,75     13.png  -
# 1 Камера в светильнике    Не включать     Светильник хирургический с видеокамерой             UTP 4x2x0,52 cat 5e/RG-6U (75 Ом)       19.png  -
# 1 Щит автоматики ОВ   Не включать     Щит автоматики системы вентиляции               ((50))UTP 4x2x0,52 cat 5e/FTP 4x2x0,52 cat 5e/ШВВП 2х0,75       3.png   -
# 1 Щит силовой Не включать     Щит электрический силовой с блоком контроля изоляции                ((50))UTP 4x2x0,52 cat 5e/ШВВП 2х0,75/ШВВП 2х0,75/UTP 4x2x0,52 cat 5e/UTP 4x2x0,52 cat 5e/ШВВП 2х0,75/ШВВП 2х0,75       2.png   -"""


# names = {}
# for line in src.split('\n'):
#   names[line.split('\t')[1]] = line.split('\t')

# t = Transaction(doc, 'Test')

# t.Start()

# for el in els:
#   name = el.LookupParameter('Имя типа').AsString()
#   if name in names:
#       el.LookupParameter('Описание').Set(names[name][2]),
#       el.LookupParameter('Комментарии к типоразмеру').Set(names[name][3]),
#       el.LookupParameter('URL').Set(names[name][4]),
#       el.LookupParameter('Изготовитель').Set(names[name][5]),
#       el.LookupParameter('Группа модели').Set(names[name][6]),
#       el.LookupParameter('Ключевая пометка').Set(names[name][7]),
#       el.LookupParameter('КабельСигнальный').Set(names[name][8]),
#       el.LookupParameter('КабельСиловой').Set(names[name][9]),
#   else:
#       print(name)

# t.Commit()




# schedules = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Schedules).WhereElementIsNotElementType().ToElements()

# schedules = list(filter(lambda x: 'Пом ' in x.Name, schedules))

# # print(schedules)

# for sch in sorted(schedules, key=lambda x: x.Name):
#   print(sch.Name)
#   for sch_filter in sch.Definition.GetFilters():
#       print(sch_filter.FieldId)
#       print(sch_filter.FilterType)
#       if sch_filter.IsDoubleValue:
#           print(sch_filter.GetDoubleValue())
#       elif sch_filter.IsElementIdValue:
#           print(sch_filter.GetElementIdValue())
#       elif sch_filter.IsIntegerValue:
#           print(sch_filter.GetIntegerValue())
#       elif sch_filter.IsStringValue:
#           print(sch_filter.GetStringValue())
#       elif sch_filter.IsNullValue:
#           print('Null')
#       else:
#           print('--------------else')
#   print()

    # sch_filter = sch.Definition.GetFilters()[0]
    # sch_filter.SetValue(sch.Name)
    # sch.Definition.SetFilter(0, sch_filter)


# print(schedule.Definition.GetFilters()[0].GetStringValue())
# filter = schedule.Definition.GetFilters()[0]
# filter.SetValue('111')
# schedule.Definition.SetFilter(0,filter)
# print(schedule.Definition.GetFilters()[0].GetStringValue())

# schedule_filter = schedule.Definition.GetFilters()[0]
# print(schedule_filter.FieldId)






# # sel = [doc.GetElement(elid) for elid in uidoc.Selection.GetElementIds()]

# # els = FilteredElementCollector(doc, ElementId(2446)).ToElements()

# # schedule = sel[0].GetTableData().GetSectionData(SectionType.Body)
# # schedule.InsertRow(schedule.FirstRowNumber + 1)
# # schedule.SetCellText(0, 0, '123')

# # for el in els:
# #     el.LookupParameter('й1').Set('1q')

# # print(els)


# # data = schedule.GetTableData().GetSectionData(SectionType.Body)

# # # for i in range(data.LastRowNumber - 0):
# # #   data.RemoveRow(i+1)

# # data.InsertRow(0)

# # # print(data.LastRowNumber)





# src = '''ОК-8   2100    900     ГОСТ 30674-99   ОП Б2 2100-900 (4М-12-4М-20-И4) 0
# ОК-40   2100    2700        ГОСТ 30674-99   ОП Б2 2100-2700 (4М-12-4М-20-И4)    Глухое
# ОК-37   2100    2700        ГОСТ 30674-99   ОП Б2 2100-2700 (4М-12-4М-12-И4)    0
# ОК-31   2100    1800        ГОСТ 30674-99   ОП Б2 2100-1800 (4М-12-4М-20-И4)    Глухое
# ОК-30   2100    1800        ГОСТ 30674-99   ОП Б2 2100-1800 (4М-12-4М-20-И4)    Глухое
# ОК-26   2100    1800        ГОСТ 30674-99   ОП Б2 2100-1800 (4М-12-4М-12-И4)    0
# ОК-14   2100    900     ГОСТ 30674-99   ОП Б2 2100-900 (4М-12-4М-12-И4) 0
# ОК-10   2100    900     ГОСТ 30674-99   ОП Б2 2100-900  (4М-12-4М-20-И4)    Глухое
# Ов-9    1310    1100        ТУ производителя    1310х1100 (h) Rg    внутреннее окно, однокамерный стеклопакет, рентгензащитное Pb=1мм, поставка Philips
# ОК-23   2100    1500        ГОСТ 30674-99   ОП Б2 2100-1500 (4М-12-4М-20-И4)    Глухое
# ОК-11   2100    900     ГОСТ 30674-99   ОП Б2 2100-900  (4М-12-4М-20-И4)    Глухое'''

# raws = src.split('\n')

# # els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Doors).WhereElementIsElementType().ToElements()

# # sel = [doc.GetElement(elid) for elid in uidoc.Selection.GetElementIds()]

# # sel_symbol = doc.GetElement(sel[0].GetTypeId())

# # for raw in raws:
# #   raw = raw.split('\t')
# #   # if float(raw[2]) >= 1400:
# #   if float(raw[2]) < 1400:
# #       symbols = list(filter(lambda x: x.LookupParameter('Имя типа').AsString() == raw[0], els))
# #       if symbols:
# #           symbol = symbols[0]
# #       else:
# #           symbol = sel_symbol
# #           symbol = symbol.Duplicate(raw[0])
# #       symbol.LookupParameter('Наличие окна').Set(1500 / k if raw[1] == 'окно' else 15000 / k)
# #       symbol.LookupParameter('Ширина').Set(float(raw[2]) / k)
# #       symbol.LookupParameter('Комментарии к типоразмеру').Set(raw[6])
# #       symbol.LookupParameter('ADSK_Марка').Set(raw[0])
# #       symbol.LookupParameter('ADSK_Наименование').Set(raw[5])
# #       symbol.LookupParameter('ADSK_Обозначение').Set(raw[4])
# #       if symbol.LookupParameter('Левая'):
# #           symbol.LookupParameter('Левая').Set(True if raw[3] == 'Л' else False)





# # els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Windows).WhereElementIsElementType().ToElements()

# # sel = [doc.GetElement(elid) for elid in uidoc.Selection.GetElementIds()]

# # sel_symbol = doc.GetElement(sel[0].GetTypeId())

# # for raw in raws:
# #   raw = raw.split('\t')
# #   symbols = list(filter(lambda x: x.LookupParameter('Имя типа').AsString() == raw[0], els))
# #   if symbols:
# #       symbol = symbols[0]
# #   else:
# #       symbol = sel_symbol
# #       symbol = symbol.Duplicate(raw[0])
# #   symbol.LookupParameter('Примерная ширина').Set((float(raw[2]) - 40) / k)
# #   symbol.LookupParameter('Примерная высота').Set((float(raw[1]) - 50) / k)
# #   symbol.LookupParameter('Комментарии к типоразмеру').Set(raw[6] if raw[6] != '0' else '')
# #   symbol.LookupParameter('ADSK_Марка').Set(raw[0])
# #   symbol.LookupParameter('ADSK_Наименование').Set(raw[5])
# #   symbol.LookupParameter('ADSK_Обозначение').Set(raw[4])

















# src = '''0.701    Лаборатория биоинформатики  27,75       5   29,14
# 0.702 Переговорная    29,53       5   31,01
# 0.703 Комната персонала   20,88       5   21,92
# 0.704 Кабинет заведующего 13,87       5   14,56
# 0.704a    Тамбур  3,2     5   3,36
# 0.705 Зона приема биоматериалов   19,68       5   20,66
# 0.706 Лаборатория ПГД 31,9        5   33,50
# 0.706a    Шлюз    5,55        5   5,83
# 0.707 Цитогенетическая лаборатория    34,97       5   36,72
# 0.707a    Шлюз    5,65        5   5,93
# 0.708 Лаборатория клеточных культур   25,84   1   15  29,72
# 0.708a    Шлюз    4,58        5   4,81
# 0.709 Лаборатория клеточной микроскопии   43,48       5   45,65
# 0.710 Лаборатория общего секвенирования   32,25   1   15  37,09
# 0.711 Лаборатория Pre-PCR 20,68   1   15  23,78
# 0.712 Лаборатория RT-PCR  14,15   1   15  16,27
# 0.713 Лаборатория PCR 9,35    1   15  10,75
# 0.714 Молекулярно-генетическая лаборатория    41,26   1   15  47,45
# 0.714a    Шлюз    3,64        5   3,82
# 0.715 Лаборатория электрофоретических исследований    19,31   1   15  22,21
# 0.715a    Шлюз    4,54        5   4,77
# 0.716 Лаборатория микрофлюидики   23,87       5   25,06
# 0.717 Зона подготовки проб    43,28       5   45,44
# 0.718 Лаборатория масс-спектрометрии  46,65       5   48,98
# 0.718а    Техническое помещение   3,13        5   3,29
# 0.720 Материальная    24,37       5   25,59
# 0.721 Помещение для временного хранения отходов   6       5   6,30
# 0.722 Помещение для приготовления растворов, чистой воды  15  1   15  17,25
# 0.723 Стерилизационная    12,26   1   15  14,10
# 0.724 Криохранилище   12,6        5   13,23
# 0.725 Помещение уборочного инвентаря  4       5   4,20
# 0.726.1   Санпропускник 1 7,46        5   7,83
# 0.726.2   Душевая 6,96        5   7,31
# 0.726.3   Санпропускник 2 7,15        5   7,51
# 0.727 Туалет  2,11        5   2,22
# 0.728 Коридор 173,71      5   182,40
# 0.792 Туалет  2,03        5   2,13
# 0.793 Туалет  2,16        5   2,27
# 0.029 Помещение криохранилища (-80 градусов Цельсия)  45,21   1   15  51,99
# 0.030 Помещение криохранилища (-196 градусов Цельсия) 26,86       5   28,20
# 0.031 Лаборатория аликвотирования 39,75   1   15  45,71
# 0.031a    Шлюз    4,02        5   4,22
# 0.032 Автоклавная 4,79        5   5,03
# 0.033 Техническое помещение   2,42        5   2,54
# 0.034 Упаковка чистых инструментов    7,25        5   7,61
# 0.035 Лаборатория геномного секвенирования    45,27       5   47,53
# 0.036 Кабинет научных сотрудников 25,74       5   27,03
# 0.037 Зона для переодевания персонала с санузлом  32,76       5   34,40
# 0.038 Помещение хранения специальной одежды   9,61        5   10,09
# 0.039 Материальная    15      5   15,75
# 0.040 Временное хранение отходов  6,02        5   6,32
# 0.041 Помещение уборочного инвентаря  5,8     5   6,09
# 0.042 Коридор 91,76       5   96,35'''

# rows = src.split('\n')


# els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()
# for index, el in enumerate(els):
#   row = rows[index].split('\t')
#   # if row[0] == el.LookupParameter('Номер').AsString():
#   #     number = row[3]
#   el.LookupParameter('Номер').Set(row[0])
#   el.LookupParameter('Имя').Set(row[1])
#   el.LookupParameter('Старая площадь число').Set(float(row[2].replace(',', '.')))
#   el.LookupParameter('Целевая площадь число').Set(float(row[5].replace(',', '.')))



# t.Commit()

# main()
