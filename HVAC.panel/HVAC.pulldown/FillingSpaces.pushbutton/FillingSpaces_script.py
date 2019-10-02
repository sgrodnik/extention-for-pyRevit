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
src = """1	Рентгеноперационная	Б		21	50,4	2,6	131			3000	2400	FFD 12/14	Вытяжные шкафы	П1	В1
2	Предоперационная	В		21	12,6	2,6	33			300	200	H11	4АПН+ЗКСД	П1	В1
3	Санпропускник	Г		21	16,3	2,6	42			250		ДПУ		П1	.
3.1	Туалет	Г		21	1,9	2,6	5				50		ДПУ		В2
3.2	Душ	Г		21	1,9	2,6	5				75		ДПУ		В2
3.3	Коридор	Г		21	4,7	2,6	12				100		ДПУ		В1
4	Комната управления	В		21	10,8	2,6	28			250	300	H11	4АПН+ЗКСД	П1	В1
5	Наркозная	В		21	11,1	2,6	29			800	800	H11	Вытяжной шкаф	П1	В1
6	Шлюз-перекладчик	Г		21	13,2	2,6	34			300	200	H11	4АПН+ЗКСД	П1	В1
7	Помещение хозинвентаря и дезсредств	Г		21	4,2	2,6	11				50		ДПУ		В1
8	Инструментально-материальная	Г		21	7,5	2,6	20	6	4	120	80	ДПУ	ДПУ	П1	В1
9	Моечная	Г		21	9,7	2,6	25	3	4	80	100	ДПУ	ДПУ	П1	В1
10	Шлюз чистый	Г		21	5	2,6	13			150		ДПУ		П1	.
11	Шлюз грязный	Г		21	4,9	2,6	13				250		ДПУ		В1
к1	Общебольничный коридор			21							400		ДПУ		В1
к2	Общебольничный коридор			21							200		ДПУ		В1"""

spaces = FilteredElementCollector(doc)\
    .OfCategory(BuiltInCategory.OST_MEPSpaces)\
    .WhereElementIsNotElementType().ToElements()

# values = [x.split('\t') for x in src.split('\n')] Практика мэппинга:
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
categ = map(lambda x: x[2], values)
supplyAir = map(lambda x: x[10], values)
exhaustAir = map(lambda x: x[11], values)
pritok = map(lambda x: x[12], values)
# for i, pos in enumerate(values):
# 	print '------------------------------------'
# 	print i
# 	print pos
vityajka = map(lambda x: x[13], values)
sysP = map(lambda x: x[14], values)
sysV = map(lambda x: x[15], values)
# vityajka = [i[13] for i in values]
# vityajka = []
# for i in values:
# 	print('{} - {}'.format(i[12], i[13]))
# 	vityajka.append(i[13])
# print 1222
existNums = []
for sp in sorted(spaces, key=lambda x: x.LookupParameter('Номер').AsString()):
    n = sp.LookupParameter('Номер').AsString()
    existNums.append(n)
    if n in nums:
        print(n)
        index = nums.index(sp.LookupParameter('Номер').AsString())
        sA = supplyAir[index]
        eA = exhaustAir[index]
        s = ''
        s1 = ''
        s2 = ''
        # if sA:
        #     s1 = '+{}'.format(sA)
        # if eA:
        #     s2 = '-{}'.format(eA)
        if sA and sA != '0':
            s1 = '{}: +{}'.format(sysP[index], sA)
        if eA and eA != '0':
            s2 = '{}: -{}'.format(sysV[index], eA)
        if s1 and s2:
            s = s1 + ', ' + s2
        else:
            s = (s1 or s2) or s
        #sp.LookupParameter('Комментарии').Set(s + ', ' + pritok[index] + ', ' + vityajka[index])
        s += '\n{}, {}, {}'.format(categ[index], pritok[index], vityajka[index]).replace(' ,', '').replace('\n, ', '\n')  # Доп инфа
        if s[-1] == ' ':
        	s = s[:-1]
        if s[-1] == ',':
        	s = s[:-1]
        sp.LookupParameter('Комментарии').Set(s)
        sp.LookupParameter('ADSK_Наименование приточной системы').Set(sysP[index])
        sp.LookupParameter('ADSK_Наименование вытяжной системы').Set(sysV[index])
        # sp.LookupParameter('Фильтр').Set(pritok[index])
        # sp.LookupParameter('Вытяжка').Set(vityajka[index])
        #if sA: s.LookupParameter('ADSK_Расчетный приток').Set(int(sA))
        # elif not(s.LookupParameter('ADSK_Расчетный приток').AsDouble()): s.LookupParameter('ADSK_Расчетный приток').Set(sA)
        #if eA: s.LookupParameter('ADSK_Расчетная вытяжка').Set(eA)
        # elif not(s.LookupParameter('ADSK_Расчетная вытяжка').AsDouble()): s.LookupParameter('ADSK_Расчетная вытяжка').Set(eA)
    else:
        print('В таблице нет {} {}'.format(
            n, sp.LookupParameter('Имя').AsString()))

print('--------------------------------------------------------')

for sp in sorted(list(set(nums))):
    if sp not in existNums:
        print('В модели нет {}'.format(sp))
t.Commit()
