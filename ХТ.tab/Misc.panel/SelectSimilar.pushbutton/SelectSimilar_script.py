# -*- coding: utf-8 -*-
"""Выбрать подобные элементы - все экземпляры семейств выбранных элементов"""
__title__ = 'Выбрать\nподобные'
__author__ = 'SG'

import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *
from Autodesk.Revit.DB import FamilyInstanceFilter, FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

sel = uidoc.Selection.GetElementIds()
if len(sel) != 0:
    string = ''
    arr = []
    for id in sel:
        el = doc.GetElement(id)
        familyType = doc.GetElement(el.GetTypeId())
        familyName = familyType.FamilyName

        if familyType.GetType().Name != 'FamilySymbol':
            col = FilteredElementCollector(doc).OfClass(el.GetType()).WhereElementIsNotElementType().ToElements()
            els = []
            info = ''
            namesList = []
            countList = []
            for i in col:
                if doc.GetElement(i.GetTypeId()).FamilyName == familyName:
                    els.append(i.Id)
                    familyTypeName = i.Name
                    if familyTypeName not in namesList:
                        namesList.append(familyTypeName)
                        countList.append(1)
                    else:
                        countList[namesList.index(familyTypeName)] += 1
            for n, c in zip(namesList, countList):
                info += '{}: {} экземпляров\n'.format(n, c)
            arr.extend(els)
            counter = len(els)

        else:
            familySymbolIds = el.Symbol.Family.GetFamilySymbolIds()
            counter = 0
            els = []
            info = ''
            for i in list(familySymbolIds):
                col = FilteredElementCollector(doc).WherePasses(FamilyInstanceFilter(doc, i)).ToElementIds()
                for j in col:
                    els.append(j)
                familyName = doc.GetElement(i).get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
                familyCount = len(col)
                counter += familyCount
                info += familyName + ((': ' + str(familyCount) + ' экземпляров\n') if familyCount > 0 else '\n')
            arr.extend(els)

        uidoc.Selection.SetElementIds(List[ElementId](arr))

        string += info
    print('Выбрано ' + str(counter) + ' элементов')
    print(string)
