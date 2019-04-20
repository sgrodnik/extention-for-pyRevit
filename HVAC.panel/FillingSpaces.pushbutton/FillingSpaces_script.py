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
src = """302.23	Палата послеоперационная на 4 места		А	21	43,3	3	130	18		2350	1880	H14	4АПН+ЗКСД
302.27	Операционная		А	21	40,4	3	121			7800	6240	FFD3030	Вытяжные шкафы
302.32	Реанимационная палата		А	21	17,4	3	52	18		950	760	H14	4АПН+ЗКСД
302.25	Эндоскопическая операционная		Б	21	38,1	3	114	15		1700	1360	H13	Вытяжные шкафы
302.29	Палата послеоперационная (изолятор)		Б	21	17,8	3	53	15		800	640	H13	4АПН+ЗКСД
302.35	Стерилмзационная		Б	21	14,5	3	44	15		650	520	H13	4АПН+ЗКСД
302.39	Наркозная		В	21	19,8	3	59			800	800	H11	Вытяжной шкаф
302.21	Шлюз		В	21	6,9	3	21			200		H11	.
302.22	Пост дежурной медсестры		В	21	10,9	3	33				200		4АПН+ЗКСД
302.24	Предоперационная		В	21	11,2	3	34			300	300	H11	4АПН+ЗКСД
302.26	Предоперационная		В	21	9	3	27			200	200	H11	4АПН+ЗКСД
302.28	Шлюз		В	21	4,6	3	14			200		H11	.
302.31	Шлюз		В	21	4,4	3	13			200		H11	.
302.20	Помещение хранения грязного белья		Г	21	5,4	3	16				100		ДПУ
302.30	С/У		Г	21	1,8	3	5				75		ДПУ
302.33	С/У		Г	21	1,2	3	4				75		ДПУ
302.34	Пост дежурной медсестры		Г	21	3,9	3	12				100		ДПУ
302.34	Слив		Г	21		3	0				100		.
302.36	Комната персонала		Г	21	23,1	3	69				200		ДПУ
302.к	Коридор		Г	21		3	0				1300		ДПУ"""

spaces = FilteredElementCollector(doc)\
.OfCategory(BuiltInCategory.OST_MEPSpaces)\
.WhereElementIsNotElementType().ToElements()

#values = [x.split('\t') for x in src.split('\n')] Практика мэппинга:
values = map(lambda x: x.split('\t'), src.split('\n'))
# for i in values:
#     print('-------------')
#     for j in i:
#       if j:
#           print(j)
#       else:
#           print('-')
nums = map(lambda x: x[0], values)
# print nums
names = map(lambda x: x[1], values)
categ = map(lambda x: x[3], values)
supplyAir = map(lambda x: x[10], values)
exhaustAir = map(lambda x: x[11], values)
pritok = map(lambda x: x[12], values)
# for i, pos in enumerate(values):
# 	print '------------------------------------'
# 	print i
# 	print pos
vityajka = map(lambda x: x[13], values)
# vityajka = [i[13] for i in values]
# vityajka = []
# for i in values:
# 	print('{} - {}'.format(i[12], i[13]))
# 	vityajka.append(i[13])
# print 1222
for sp in spaces:
    n = sp.LookupParameter('Номер').AsString()
    print n
    if n in nums:
        index = nums.index(sp.LookupParameter('Номер').AsString())
        sA = supplyAir[index]
        eA = exhaustAir[index]
        s = ''
        s1 = ''
        s2 = ''
        if sA: s1 = '+{}'.format(sA)
        if eA: s2 = '-{}'.format(eA)
        if s1 and s2: s = s1 + ', ' + s2
        else: s = (s1 or s2) or s
        #sp.LookupParameter('Комментарии').Set(s + ', ' + pritok[index] + ', ' + vityajka[index])
        s += '\n{}, {}, {}'.format(categ[index], pritok[index], vityajka[index])
        sp.LookupParameter('Комментарии').Set(s)
        #sp.LookupParameter('Фильтр').Set(pritok[index])
        #sp.LookupParameter('Вытяжка').Set(vityajka[index])
        #if sA: s.LookupParameter('ADSK_Расчетный приток').Set(int(sA))
        #elif not(s.LookupParameter('ADSK_Расчетный приток').AsDouble()): s.LookupParameter('ADSK_Расчетный приток').Set(sA)
        #if eA: s.LookupParameter('ADSK_Расчетная вытяжка').Set(eA)
        #elif not(s.LookupParameter('ADSK_Расчетная вытяжка').AsDouble()): s.LookupParameter('ADSK_Расчетная вытяжка').Set(eA)
    else:
        print('Нету {} {}'.format(n, sp.LookupParameter('Имя').AsString()))

t.Commit()