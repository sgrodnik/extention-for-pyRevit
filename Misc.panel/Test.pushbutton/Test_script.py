# -*- coding: utf-8 -*-
""""""
__title__ = 'Test'
__author__ = 'SG'

import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *
from Autodesk.Revit.DB import SectionType, ElementId, PartUtils, ViewOrientation3D, XYZ, FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, Line, Structure, ParameterFilterElement, ParameterFilterRuleFactory, ElementParameterFilter, ViewDuplicateOption
import sys
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
# from Autodesk.Revit.ApplicationServices.Application import Create
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application
k = 304.8


from pyrevit import script
output = script.get_output()

t = Transaction(doc, 'Test')
t.Start()
els = FilteredElementCollector(doc).WhereElementIsElementType().ToElements()
for el in els:
    if el.LookupParameter('Изготовитель'):
        if el.LookupParameter('Изготовитель').HasValue:
            old = el.LookupParameter('Изготовитель').AsString() == 'Лиссант'
            if old:
                el.LookupParameter('Изготовитель').Set('Россия')
                print('{} {}'.format(old, el.LookupParameter('Имя типа').AsString()))
t.Commit()


# sel = [doc.GetElement(elid) for elid in uidoc.Selection.GetElementIds()]

# cir0 = [i for i in sel if i.Name == 'Линии детализации'][0]
# cir1 = [i for i in sel if i.Name == 'Линии детализации'][1]
# # view = [i for i in sel if i.Name != 'Линии детализации'][0]

# t = Transaction(doc, 'Test')
# t.Start()

# # for i in dir(view.Location):
# # 	print(i)
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

