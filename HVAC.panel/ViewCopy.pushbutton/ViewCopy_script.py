# -*- coding: utf-8 -*-
""""""
__title__ = 'Копирование\nвида'
__author__ = 'SG'

import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *

from Autodesk.Revit.DB import ParameterFilterElement, ViewDuplicateOption, FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

t = Transaction(doc, 'Копирование вида')
t.Start()

# for i in dir(uidoc):
#     print(i)

viewNames = ['К02',
             'К03',
             'К04',
             'К05',
             'К06',
             'К07',
             'К08',
             'В01',
             'В02',
             'В03',
             'В04',
             'В05',
             'В06',
             'В07',
             'В08',
             'В09',
             'В10',
             'В11',
             'В12',
             'В13',
             'В14',
             'В15',
             'В16']
for name in viewNames:
    newViewId = uidoc.ActiveView.Duplicate(ViewDuplicateOption.Duplicate)
    newView = doc.GetElement(newViewId)
    newView.Name = name

# newView.RemoveFilter(oldFilterId)
# newView.AddFilter(newFilterId)
# newView.SetFilterVisibility(newFilterId, False)
# newView.Name = newName

filters = [doc.GetElement(id) for id in uidoc.ActiveView.GetFilters()]

catNames = ['Арматура воздуховодов',
            'Арматура трубопроводов',
            'Воздуховоды',
            'Воздухораспределители',
            'Гибкие воздуховоды',
            'Гибкие трубы',
            'Материалы изоляции воздуховодов',
            'Материалы изоляции труб',
            'Оборудование',
            'Соединительные детали трубопроводов',
            'Соединительные детали воздуховодов',
            'Трубы']


categoryIds = []
for cat in doc.Settings.Categories:
    if cat.Name in catNames:
        categoryIds.append(cat.Id)

# pfe = ParameterFilterElement.Create(doc, 'String', List[ElementId](categoryIds))
# pfe.SetElementFilter

# print(filters[1].Name)
# for i in dir(filters[1].GetElementFilterParameters()):
#     print(i)

t.Commit()
