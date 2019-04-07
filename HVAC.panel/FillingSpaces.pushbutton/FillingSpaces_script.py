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

src = """701.51	Ангиографическая операционная		А	21	63,9	3	192	18		3450	2760	H14	Вытяжные шкафы	П01	В01
701.50	Эндоскопическая операционная		А	21	43,4	3	130			4500	3600	FFD2024	Вытяжные шкафы	П02	В02
701.33	Операционная пластической хирургии		А	21	41,6	3	125			7800	6240	FFD3030	Вытяжные шкафы	П03	В03
701.31	Ортопедическая операционная		А	21	41,2	3	124			7800	6240	FFD3030	Вытяжные шкафы	П04	В04
701.24	Абдоминальная операционная		А	21	41,2	3	124			7800	6240	FFD3030	Вытяжные шкафы	П05	В05
701.23	Общехирургическая операционная		А	21	41,1	3	123			7800	6240	FFD3030	Вытяжные шкафы	П06	В06
701.19	Офтальмологическая операционная		А	21	45,5	3	137	18		2450	1960	H14	Вытяжные шкафы	П07	В07
701.18	Наркозная		В	21	13,3	3	40			800	800	H11	Вытяжной шкаф	П12	В12
701.26	Наркозная		В	21	13,3	3	40			800	800	H11	Вытяжной шкаф	П12	В12
701.30	Наркозная		В	21	13,1	3	39			800	800	H11	Вытяжной шкаф	П12	В12
701.34	Наркозная		В	21	14,1	3	42			800	800	H11	Вытяжной шкаф	П12	В12
701.53	Наркозная		В	21	27,7	3	83			800	800	H11	Вытяжной шкаф	П12	В12
701.21	Предоперационная		В	21	11,2	3	34			300	300	H11	4АПН+ЗКСД	П12	В12
701.25	Предоперационная		В	21	14,5	3	44			300	300	H11	4АПН+ЗКСД	П12	В12
701.32	Предоперационная		В	21	18,9	3	57			350	350	H11	4АПН+ЗКСД	П12	В12
701.46	Предоперационная		В	21	18,4	3	55			350	350	H11	4АПН+ЗКСД	П12	В12
701.48	Кабинет врача		Г	21	21,3	3	64		3		200		ДПУ	П12	В12
701.35	Комната для хранения отходов		Г	21	6,2	3	19				50		ДПУ	П12	В12
701.54	Комната персонала		Г	21	39,3	3	118		3		350		ДПУ	П12	В12
701.29	Помещение хранения медикаментов		Г	21	11	3	33				50		ДПУ	П12	В12
701.55	Помещение хранения крови		Г	21	12,6	3	38				50		ДПУ	П12	В12
701.47	Помещение		Г	21	20,3	3	61			3500	3500	THLabor2416	Вытяжные шкафы	П12	В12
702.33	Коридор		Г	21	202,7	4	811			3500	8000	THLabor2416	4АПН+ЗКСД, Вытяжные шкафы		В"""

spaces = FilteredElementCollector(doc)\
.OfCategory(BuiltInCategory.OST_MEPSpaces)\
.WhereElementIsNotElementType().ToElements()

#values = [x.split('\t') for x in src.split('\n')] Практика мэппинга:
values = map(lambda x: x.split('\t'), src.split('\n'))
# for i in values:
# 	print('-------------')
# 	for j in i:
# 		if j:
# 			print(j)
# 		else:
# 			print('-')
nums = map(lambda x: x[0], values)
names = map(lambda x: x[1], values)
categ = map(lambda x: x[3], values)
supplyAir = map(lambda x: x[10], values)
exhaustAir = map(lambda x: x[11], values)
pritok = map(lambda x: x[12], values)
vityajka = map(lambda x: x[13], values)

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
		#sp.LookupParameter('Комментарии').Set(s + ', ' + pritok[index] + ', ' + vityajka[index])
		s += '\n{}, {}, {}'.format(categ[index], pritok[index], vityajka[index])
		sp.LookupParameter('Комментарии').Set(s)
		#sp.LookupParameter('Фильтр').Set(pritok[index])
		#sp.LookupParameter('Вытяжка').Set(vityajka[index])
		#if sA: s.LookupParameter('ADSK_Расчетный приток').Set(int(sA))
		#elif not(s.LookupParameter('ADSK_Расчетный приток').AsDouble()): s.LookupParameter('ADSK_Расчетный приток').Set(sA)
		#if eA: s.LookupParameter('ADSK_Расчетная вытяжка').Set(eA)
		#elif not(s.LookupParameter('ADSK_Расчетная вытяжка').AsDouble()): s.LookupParameter('ADSK_Расчетная вытяжка').Set(eA)
	# else:
	# 	print('Нету {} {}'.format(n, sp.LookupParameter('Имя').AsString()));

t.Commit()