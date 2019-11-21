# -*- coding: utf-8 -*-
"""Расчёт длины трубопроводной петли"""
__title__ = 'Длина\nпетли'
__author__ = 'SG'


import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


k = 304.8

global handledElems
global counter
counter = 0
global OUT


def toList(iEnumerator):
    list = []
    while iEnumerator.MoveNext():
        list.append(iEnumerator.Current)
    return list


def getConnectors(elem):
    global tabs
    tabs += 1
    connectors = []
    if elem.Category.Name == 'Трубы':
        connectors = toList(elem.ConnectorManager.Connectors.GetEnumerator())
    elif elem.Category.Name == 'Соединительные детали трубопроводов'\
         and (elem.Name == 'Дренаж' or elem.Name == 'Прямая'):
        connectors = toList(
            elem.MEPModel.ConnectorManager.Connectors.GetEnumerator())
    tabs -= 1
    return connectors


def getSiblings(elem):
    global handledElems
    global counter
    global tabs
    connectorSet = getConnectors(elem)
    siblings = []
    if elem.Name == 'Дренаж' or 'Труба' in elem.Name or 'Прямая' in elem.Name:
        siblings = [elem]
    for i, connector in enumerate(connectorSet):
        counter += 1
        tabs += 1
        if connector.IsConnected:
            for con in connector.AllRefs:
                if con.Owner.Id.ToString() != connector.Owner.Id.ToString() and\
                   con.Owner.Category.Name != 'Трубопроводные системы':
                    linkedCon = con
        else:
            tabs -= 1
            continue
        owner = linkedCon.Owner
        if owner.Id.ToString() not in handledElems:
            handledElems.append(owner.Id.ToString())
            sibs = getSiblings(owner)
            if sibs:
                siblings.extend(sibs)
    tabs -= 1
    return siblings


def main():
    pipes = FilteredElementCollector(doc)\
        .OfCategory(BuiltInCategory.OST_PipeCurves)\
        .WhereElementIsNotElementType().ToElements()
    pipes = [pipe for pipe in pipes
             if '{}'.format(pipe.LookupParameter('Длина').AsDouble() * k) !=
             pipe.LookupParameter('Длина старая').AsString()
             ]

    tg = TransactionGroup(doc, "Длина петли")
    tg.Start()
    t = Transaction(doc, "blabla")
    t.Start()

    [pipe.LookupParameter('Длина старая').Set('{}'.format(pipe.LookupParameter('Длина').AsDouble() * k)) for pipe in pipes]

    global handledElems
    handledElems = []
    global tabs
    global OUT
    siblingGroups = []
    for pipe in pipes:
        siblings = []
        if pipe.Id.ToString() not in handledElems:
            handledElems.append(pipe.Id.ToString())
            tabs = 1
            siblings = getSiblings(pipe)
            siblingGroups.append(siblings)

    for group in siblingGroups:
        s = sum([e.LookupParameter('Длина').AsDouble() * k for e in group if e.LookupParameter('Длина')])
        for el in group:
            el.LookupParameter('Длина петли').Set('{:.0f}'.format(s / 1000).replace('.', ','))

    t.Commit()
    tg.Assimilate()


with open('C:\\Users\\sgrodnik\\Desktop\\so.txt', 'w', buffering=1) as o:
    # sys.stdout = o
    main()
