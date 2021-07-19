# -*- coding: utf-8 -*-
"""Добавляет к выбору цепи выбранного оборудования. Отфильтровывает из выбора лишние категории"""
__title__ = 'Выбрать цепи\nоборудования'
__author__ = 'SG'

import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *
from Autodesk.Revit.DB import ElementId, PartUtils, ViewOrientation3D, XYZ, FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter
import sys
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


sel = [doc.GetElement(elid) for elid in uidoc.Selection.GetElementIds()]

# sel = filter(lambda x: x.LookupParameter('Категория').AsValueString() == 'Электрооборудование', sel)
sel = [el for el in sel if el.LookupParameter('Категория').AsValueString() == 'Электрооборудование']

if sel:
	sel_save = [el.Id for el in sel if 'ейк' not in el.LookupParameter('Тип').AsValueString()]

	system_ids_to_select = []
	for el in sel:
		# print(el.Id)
		if el.MEPModel.ElectricalSystems:
			for system in el.MEPModel.ElectricalSystems:
				if system.Id not in system_ids_to_select:
					system_ids_to_select.append(system.Id)
		elif 'ейк' in el.LookupParameter('Тип').AsValueString():
			target = el.LookupParameter('Цепь').AsString()
			system_ids_to_select.append(ElementId(int(target)))



	uidoc.Selection.SetElementIds(List[ElementId](system_ids_to_select + sel_save))
