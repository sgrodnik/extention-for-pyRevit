# -*- coding: utf-8 -*-
"""описание"""
__title__ = 'Материалы\n'
__author__ = 'SG'
import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *
from Autodesk.Revit.DB import Material, ParameterType, FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId
import sys
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# names = {'Аллюминий': 'Алюминий',
# 'Алюминий': 'Алюминий',
# 'Белый пластик': 'Пластик белый',
# 'Голубая подсветка': 'Голубая подсветка',
# 'Голубое стекло': 'Стекло голубое',
# 'Дисплей': 'Дисплей',
# 'Дисплей красный': 'Дисплей красный',
# 'Желтый пластик': 'Пластик желтый',
# 'ЖК экран серый': 'ЖК экран серый',
# 'Зеленое стекло': 'Стекло зеленое',
# 'Зеленый пластик': 'Пластик зеленый',
# 'Кнопки': 'Кнопки',
# 'Кнопки синие': 'Кнопки синие',
# 'Колеса пластик': 'Колеса пластик',
# 'Колеса резина': 'Колеса резина',
# 'Колесо': 'Колесо',
# 'Корпус': 'Корпус',
# 'Корпус2': 'Корпус2',
# 'Красное стекло': 'Стекло красное',
# 'Красный пластик': 'Пластик красный',
# 'Логотип': 'Логотип',
# 'Материал': 'Материал',
# 'Материал 2': 'Материал 2',
# 'Материал 3': 'Материал 3',
# 'Материал дверей': 'Материал дверей',
# 'Материал дисплея': 'Материал дисплея',
# 'Материал корпуса': 'Материал корпуса',
# 'Материал ручек': 'Материал ручек',
# 'Материал трубы': 'Материал трубы',
# 'Металл': 'Металл',
# 'Мокрый асфальт': 'Мокрый асфальт',
# 'Монитор': 'Монитор',
# 'Нержавеющая сталь': 'Нержавеющая сталь',
# 'Нержавеющая сталь, красная': 'Нержавеющая сталь, красная',
# 'Нержавеющая сталь, серебристая': 'Нержавеющая сталь',
# 'Ножка': 'Ножка',
# 'Ножки': 'Ножки',
# 'Ножки черные, резина': 'Ножки черные, резина',
# 'Оцинковка': 'Оцинковка',
# 'Пластик бежевый': 'Пластик бежевый',
# 'Пластик белый': 'Пластик белый',
# 'Пластик белый (светлобежевый)': 'Пластик белый (светлобежевый)',
# 'Пластик голубой': 'Пластик голубой',
# 'Пластик желтый': 'Пластик желтый',
# 'Пластик зеленый': 'Пластик зеленый',
# 'Пластик зеленый светлый': 'Пластик зеленый светлый',
# 'Пластик зеленый темный': 'Пластик зеленый темный',
# 'Пластик коричневый': 'Пластик коричневый',
# 'Пластик красный': 'Пластик красный',
# 'Пластик светлобежевый': 'Пластик светлобежевый',
# 'Пластик серый': 'Пластик серый',
# 'Пластик серый светлый': 'Пластик серый светлый',
# 'Пластик серый темный': 'Пластик серый темный',
# 'Пластик синий': 'Пластик синий',
# 'Пластик синий светлый': 'Пластик синий светлый',
# 'Пластик черный': 'Пластик черный',
# 'Прозрачное стекло': 'Стекло',
# 'Резина': 'Резина',
# 'Резина черная': 'Резина',
# 'Ручка черная, пластик': 'Пластик черный',
# 'Светло-серый пластик': 'Пластик серый светлый',
# 'Сенсорный экран, черный': 'Сенсорный экран, черный',
# 'Серый пластик': 'Пластик серый',
# 'Синее светящееся стекло': 'Синее светящееся стекло',
# 'Синий пластик': 'Пластик синий',
# 'Сталь': 'Сталь',
# 'Сталь золотая': 'Сталь золотая',
# 'Сталь нержавеющая': 'Сталь нержавеющая',
# 'Стекло': 'Стекло',
# 'Стекло цвета морской волны': 'Стекло цвета морской волны',
# 'Стекло черное': 'Стекло черное',
# 'Столешница серая': 'Столешница серая',
# 'Темное стекло': 'Стекло темное',
# 'Хром': 'Хром',
# 'Черная порошковая краска': 'Черная порошковая краска',
# 'Черное стекло': 'Стекло черное',
# 'Черный пластик': 'Пластик черный',
# 'Черный экран': 'Черный экран',
# 'Чугун': 'Чугун',
# }

names = {'Пластик фиолетовый': 'Пластик фиолетовый',
'Стекло оранжевое': 'Стекло оранжевое',
'Аллюминий': 'Алюминий',
'Неон голубой': 'Голубая подсветка',
'Пластик бежевый': 'Пластик бежевый',
'Белая эмаль': 'Пластик белый',
'Плакстик белый': 'Пластик белый',
'Плакстик белый': 'Пластик белый',
'Пластик белый': 'Пластик белый',
'Пластик белый': 'Пластик белый',
'Пластик белый': 'Пластик белый',
'Пластик белый': 'Пластик белый',
'Пластик белый': 'Пластик белый',
'Пластик белый': 'Пластик белый',
'Пластик белый': 'Пластик белый',
'Пластик белый': 'Пластик белый',
'Пластик белый': 'Пластик белый',
'Пластик белый': 'Пластик белый',
'Пластик белый': 'Пластик белый',
'Пластик белый': 'Пластик белый',
'Пластик желтый': 'Пластик желтый',
'Зеленый монитор': 'Пластик зеленый',
'Пластик зеленый': 'Пластик зеленый',
'Пластик зеленый': 'Пластик зеленый',
'Пластик зеленый': 'Пластик зеленый',
'Пластик зеленый': 'Пластик зеленый',
'Палстик темно-зеленый': 'Пластик зеленый темный',
'Красный пластик': 'Пластик красный',
'Красный пластик': 'Пластик красный',
'Пластик красный': 'Пластик красный',
'Пластик красный': 'Пластик красный',
'Пластик красный': 'Пластик красный',
'Пластик красный': 'Пластик красный',
'Пластик красный': 'Пластик красный',
'Пластик красный': 'Пластик красный',
'Пластик красный': 'Пластик красный',
'Пластик красный': 'Пластик красный',
'Пластик красный': 'Пластик красный',
'Пластик светло-бежевый': 'Пластик светлобежевый',
'Пластик светло-желтый (как грязный)': 'Пластик светлобежевый',
'Пластик серый': 'Пластик серый',
'Пластик серый': 'Пластик серый',
'Пластик серый': 'Пластик серый',
'Пластик серый': 'Пластик серый',
'Серый пластик': 'Пластик серый',
'Серый пластик': 'Пластик серый',
'Пластик белый (сероватый)': 'Пластик серый светлый',
'Пластик белый (сероватый)': 'Пластик серый светлый',
'Пластик серый (почти белый)': 'Пластик серый светлый',
'Темно-серый пластик': 'Пластик серый темный',
'Кнопки синие': 'Пластик синий',
'пластик синий': 'Пластик синий',
'Пластик синий': 'Пластик синий',
'Пластик синий': 'Пластик синий',
'Пластик синий': 'Пластик синий',
'Пластик синий': 'Пластик синий',
'Пластик синий': 'Пластик синий',
'Ткань синяя': 'Пластик синий',
'Монитор черный': 'Пластик черный',
'Пластик черный': 'Пластик черный',
'пластик черный': 'Пластик черный',
'Пластик черный': 'Пластик черный',
'Пластик черный': 'Пластик черный',
'Пластик черный': 'Пластик черный',
'Пластик черный': 'Пластик черный',
'Пластик черный': 'Пластик черный',
'Пластик черный': 'Пластик черный',
'Пластик черный': 'Пластик черный',
'Пластик черный': 'Пластик черный',
'Пластик черный': 'Пластик черный',
'Пластик черный': 'Пластик черный',
'Пластик черный': 'Пластик черный',
'Пластик черный': 'Пластик черный',
'Пластик черный глянец': 'Пластик черный',
'Платик черный глянец': 'Пластик черный',
'Черный пластик': 'Пластик черный',
'Черный пластик': 'Пластик черный',
'Краска серебраянная': 'Сталь',
'Металл': 'Сталь',
'Металл': 'Сталь',
'Металл': 'Сталь',
'Металл': 'Сталь',
'Металл серебсристый': 'Сталь',
'сталь нержавеющая': 'Сталь нержавеющая',
'Сталь нержавеющая': 'Сталь нержавеющая',
'Палстик прозрачный': 'Стекло',
'Палстик прозрачный': 'Стекло',
'Палстик прозрачный': 'Стекло',
'Пластик прозрачный': 'Стекло',
'Пластик прозрачный': 'Стекло',
'Пластик прозрачный': 'Стекло',
'Пластик прозрачный': 'Стекло',
'Пластик прозрачный': 'Стекло',
'стекло': 'Стекло',
'Стекло': 'Стекло',
'Стекло': 'Стекло',
'Стекло': 'Стекло',
'Стекло': 'Стекло',
'Стекло': 'Стекло',
'Дисплей': 'Стекло черное',
'ЖК экран белый светящийся': 'Стекло черное',
'ЖК экран зеленый светящийся': 'Стекло черное',
'ЖК экран серый': 'Стекло черное',
'ЖК экран фиолетовый - светится': 'Стекло черное',
'ЖК экран черный': 'Стекло черное',
'Экран': 'Стекло черное',
'Экран': 'Стекло черное',
'Экран белый светящийся': 'Стекло черное',
'Экран черный': 'Стекло черное',
'Материал': 'Материал',
'Пластик черынй матовый': 'Чугун'
}

ids = [47945,
47987,
373122,
375375,
377461,
378800,
381142,
382522,
383894,
385122,
386324,
387562,
388767,
389970,
391929,
393079,
394229,
395434,
396765,
398040,
400988,
403286,
408378,
409619,
411852,
414828,
417170,
418519,
419856,
426477,
562883,
564375,
566480,
568623,
569964]

materials = FilteredElementCollector(doc)\
	.OfCategory(BuiltInCategory.OST_Materials)\
	.WhereElementIsNotElementType().ToElements()

# for mat in materials:
# 	if mat.Name == "My Material":
# 		orig = mat


t = Transaction(doc, 'Материалы')
t.Start()

# types = [doc.GetElement(doc.GetElement(id).GetTypeId()) for id in uidoc.Selection.GetElementIds()]

# params = []
# for type in types:
# 	for param in type.GetOrderedParameters():
# 		if param.Definition.ParameterType == ParameterType.Material:
# 			if param.Name in names.keys():
# 				for mat in materials:
# 					if mat.Name == names[param.Definition.Name]:
# 						break
# 				param.Set(mat.Id)

for id in ids:
	for param in doc.GetElement(ElementId(id)).GetOrderedParameters():
		if param.Definition.ParameterType == ParameterType.Material:
			if param.Definition.Name in names.keys():
				for mat in materials:
					if mat.Name == names[param.Definition.Name]:
						break
				param.Set(mat.Id)

# params = []
# for type in types:
# 	for param in type.GetOrderedParameters():
# 		if param.Definition.ParameterType == ParameterType.Material:
# 			# print(param.AsElementId)
# 			print('{}:::ff{}:::ff{}:::ff{}'.format(param.Definition.Name, param.AsValueString(), type.LookupParameter('Ключевая пометка').AsString(), type.Id))
# 			# print('{}:::{}'.format(param.Definition.Name, doc.GetElement(param.AsElementId).Name))

# for type in types:
# 	type.LookupParameter('Ключевая пометка').Set('Бизюгин')

# types[0].GetOrderedParameters()

# for param in params:
# 	print(param.Definition.Name)

# for name in names.split('\n'):
# 	orig.Duplicate(name)

# Material.Create(doc, "My Material")

# print(materials)



t.Commit()