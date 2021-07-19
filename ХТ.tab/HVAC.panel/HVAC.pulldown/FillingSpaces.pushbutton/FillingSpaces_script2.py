# -*- coding: utf-8 -*-
"""Описание"""
__title__ = 'Заполнение\nпространств'
__author__ = 'SG'

import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

t = Transaction(doc, 'Заполнение пространств')
t.Start()

# в конце строки не должно быть пустоты
src = """1	Лестничная клетка
2	Коридор
3	Коридор
4	Электрощитовая
5	Серверная
6	Шлюз
7	Шлюз
8	Комната психологической разгрузки
9	Проход
9.1	Проход
9.2	Проход
10	Кладовая расходников НДА
12	Кладовая инструментов, шовного материала
13	Кладовая НДА
14	Помещение подготовки пациента
15	Предоперационная
16	Помещение подготовки пациента
17	Операционная
18	Стерилизационная для экстренной стерилизации
19	Операционная
20	Помещение хранения мед. отходов
21	Шлюз
22	Шлюз для RG оборудования
23	Помещение мойки и обеззараживания НДА
24	Помещение подготовки пациента
25	Предоперационная
26	Помещение подготовки пациента
27	Операционная
28	Стерилизационная для экстренной стерилизации
29	Операционная
30	Шлюз
31	Помещение подготовки крови и кровезаминителей
32.1	Помещение
32.2	Помещение
32.3	Помещение
32.4	Помещение
32.5	Помещение
32.6	Помещение
32.7	Помещение
32.8	Помещение
32.9	Проход
33	Помещение уборочного инвентаря
34	Проход
34.1	Проход
36	Помещение уборочного инвентаря
37	Шлюз
38	Проход
39	Шлюз
40	Шлюз
41	Помещение персонала
42	Комната хранения наркотических средств
43	Проход
44	Шлюз
45	Шлюз
46	Палата  реанимации и интенсивной терапии на 7 коек
47	Манипуляционная
48	Подсобное помещение
49	Слив
50	Слив
51	Помещение персонала
52	Комната хранения чистого белья
53	Помещение старшей сестры
54	Кабинет сестры хозяйки
55	Ординаторская
56	Комната уборочного инвентаря
57	Кладовая
58	Шлюз
59	Боксированная палата
60	Слив
61	Шлюз
62	Палата интенсивной терапии
63	Санузел
64	Санузел персонала
64.1	Санузел персонала
64.2	Санузел персонала
65	Помещение подготовки инфузионных систем
69	Кладовая расходных материалов
70	Лестничная клетка"""

spaces = FilteredElementCollector(doc)\
    .OfCategory(BuiltInCategory.OST_MEPSpaces)\
    .WhereElementIsNotElementType().ToElements()

numbers = [row.split('\t')[0] for row in src.split('\n')]
names = [row.split('\t')[1] for row in src.split('\n')]

# print('{}: {}'.format(numbers, names))

for space in spaces:
	number = space.LookupParameter('Номер').AsString()
	try:
		name = names[numbers.index(number)]
	except ValueError:
		name = '--------------Ошибка--------------'
		print('{}: {}'.format(number, name))
	space.LookupParameter('Имя').Set(name)


# #values = [x.split('\t') for x in src.split('\n')] Практика мэппинга:
# values = map(lambda x: x.split('\t'), src.split('\n'))
# # for i in values:
# #     print('-------------')
# #     for j in i:
# #       if j:
# #           print(j)
# #       else:
# #           print('-')
# nums = map(lambda x: x[0], values)
# # print nums
# names = map(lambda x: x[1], values)
# # for i, pos in enumerate(values):
# #   print '------------------------------------'
# #   print i
# #   print pos
# # vityajka = [i[13] for i in values]
# # vityajka = []
# # for i in values:
# #   print('{} - {}'.format(i[12], i[13]))
# #   vityajka.append(i[13])
# # print 1222
# existNums = []
# for sp in sorted(spaces, key=lambda x: x.LookupParameter('Номер').AsString()):
# 	n = sp.LookupParameter('Номер').AsString()
# 	existNums.append(n)
# 	if n in nums:
# 		print n
# 		index = nums.index(sp.LookupParameter('Номер').AsString())
# 		sA = supplyAir[index]
# 		eA = exhaustAir[index]
# 		s = ''
# 		s1 = ''
# 		s2 = ''
# 		if sA: s1 = '+{}'.format(sA)
# 		if eA: s2 = '-{}'.format(eA)
# 		if s1 and s2: s = s1 + ', ' + s2
# 		else: s = (s1 or s2) or s
# 		#sp.LookupParameter('Комментарии').Set(s + ', ' + pritok[index] + ', ' + vityajka[index])
# #        s += '\n{}, {}, {}'.format(categ[index], pritok[index], vityajka[index])  # Доп инфа
# 		sp.LookupParameter('Комментарии').Set(s)
# 		#sp.LookupParameter('Фильтр').Set(pritok[index])
# 		#sp.LookupParameter('Вытяжка').Set(vityajka[index])
# 		#if sA: s.LookupParameter('ADSK_Расчетный приток').Set(int(sA))
# 		#elif not(s.LookupParameter('ADSK_Расчетный приток').AsDouble()): s.LookupParameter('ADSK_Расчетный приток').Set(sA)
# 		#if eA: s.LookupParameter('ADSK_Расчетная вытяжка').Set(eA)
# 		#elif not(s.LookupParameter('ADSK_Расчетная вытяжка').AsDouble()): s.LookupParameter('ADSK_Расчетная вытяжка').Set(eA)
# 	else:
# 		print('В таблице нет {} {}'.format(n, sp.LookupParameter('Имя').AsString()))

# print('--------------------------------------------------------')

# for sp in sorted(list(set(nums))):
# 	if sp not in existNums:
# 		print('В модели нет {}'.format(sp))
t.Commit()
