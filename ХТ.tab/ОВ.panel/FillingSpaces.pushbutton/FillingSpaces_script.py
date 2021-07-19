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

t = Transaction(doc, 'name')
t.Start()

src = """1.05	Б	Операционная	3200	2560	H14	Вытяжные шкафы
21	В	Наркозная	800	600	H13	Вытяжной шкаф
3	В	Предоперационная	300	250	H13	4АПН+ЗКСД
4	В	Пультов	200	300	H11	4АПН+ЗКСД
31	В	Коридор	390	310	H11	4АПН+ЗКСД
28	В	Шлюз	210		H11	
1	Г	Моечная		300		ДПУ
27	Г	Материальная		250		ДПУ
8	Г	С.П. М		200		ДПУ
14	Г	С/У		100		ДПУ
19	В	Чистый С.П.	250		H11	
20	Г	Грязный С. П.		110		ДПУ
24	Г	С.П. Ж		390		ДПУ
25	Г	Грязный С. П.		140		ДПУ
26	В	Чистый С.П.	250		H11	
29	Г	С/У		100		ДПУ
30	Г	Тамбур	200		ДПУ	"""

spaces = FilteredElementCollector(doc)\
.OfCategory(BuiltInCategory.OST_MEPSpaces)\
.WhereElementIsNotElementType().ToElements() 

#values = [x.split('\t') for x in src.split('\n')] Практика мэппинга:
values = map(lambda x: x.split('\t'), src.split('\n'))
nums = map(lambda x: x[0], values)
categ = map(lambda x: x[1], values)
names = map(lambda x: x[2], values)
supplyAir = map(lambda x: x[3], values)
exhaustAir = map(lambda x: x[4], values)
filtr = map(lambda x: x[5], values)
vityajka = map(lambda x: x[6], values)

arr = []

TransactionManager.Instance.EnsureInTransaction(doc)
for sp in spaces:
	n = sp.LookupParameter('Номер').AsString()
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
		#sp.LookupParameter('Комментарии').Set(s + ', ' + filtr[index] + ', ' + vityajka[index])
	#	s += '\n{}, {}, {}'.format(categ[index], filtr[index], vityajka[index])
		sp.LookupParameter('Комментарии').Set(s)
		#sp.LookupParameter('Фильтр').Set(filtr[index])
		#sp.LookupParameter('Вытяжка').Set(vityajka[index])
		#if sA: s.LookupParameter('ADSK_Расчетный приток').Set(int(sA))
		#elif not(s.LookupParameter('ADSK_Расчетный приток').AsDouble()): s.LookupParameter('ADSK_Расчетный приток').Set(sA)
		#if eA: s.LookupParameter('ADSK_Расчетная вытяжка').Set(eA)
		#elif not(s.LookupParameter('ADSK_Расчетная вытяжка').AsDouble()): s.LookupParameter('ADSK_Расчетная вытяжка').Set(eA)
	else:
		arr.append('Нету {} {}'.format(n, sp.LookupParameter('Имя').AsString()));

t.Commit()