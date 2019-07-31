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
src = """1	Травматологическая операционная	А		21	44,3	2,6	115	42	33	4800	3840	FFD24/24	Шкафы 500×200	П01	В01
2	Предоперационная	В		21	10,6	2,6	28	11	11	300	300	ВБД450 H11	4АПН+ЗКСД	П01	В01
3	Моечная №1	Г		21	10,4	2,6	27	0	4		100		ДПУ		В02
4	Наркозная	В		21	14,4	2,6	37	21	21	800	800	H11	Шкаф 400×150	П01	В01
5	Коридор	В		21	5	2,6	13	0	8		100		ДПУ		В02
6	Травматологическая операционная 2	А		21	42,6	2,6	111	43	35	4800	3840	FFD24/24	Шкафы 500×200	П02	В02
7	Наркозная №2	В		21	14,7	2,6	38	21	21	800	800	H11	Шкаф 400×150	П02	В02
8	Моечная №2	Г		21	9,6	2,6	25	0	4		100		ДПУ		В02
9	Предоперационная №2	В		21	15,3	2,6	40	8	8	300	300	H11	4АПН+ЗКСД	П02	В02
к	Коридор	Г		21	35	2,6	90				650		ДПУ		В01
к	Коридор	Г		21							350		ДПУ		В02"""

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
        s += '\n{}, {}, {}'.format(categ[index], pritok[index], vityajka[index])  # Доп инфа
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
