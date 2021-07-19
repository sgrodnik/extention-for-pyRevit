# -*- coding: utf-8 -*-
""""""
__title__ = 'Настил'
__author__ = 'SG'
import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *
from Autodesk.Revit.DB import IndependentTag, XYZ, FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId
import sys
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

k = 304.8

t = Transaction(doc, 'Настил')
t.Start()

sel = [doc.GetElement(id) for id in uidoc.Selection.GetElementIds()]

# for line in lines:
#   print(line.LineStyle.Name == 'Рез')

# print(sel[0].Location.Curve.GetEndPoint(0))
# print(sel[0].Location.Curve.GetEndPoint(1))

# if sel:
#     nass = []
#     lines = []
#     for el in sel:
#         nass.append(el) if el.Name == 'Настил' else None
#         lines.append(el) if el.Name == 'Линии детализации' else None

#     if len(nass) > 1:
#         print('Ошибка: выбрано несколько настилов')
#         t.Commit()
#         sys.exit()

#     nas = nass[0]

#     nas.LookupParameter('Длина линий').Set(sum([el.LookupParameter('Длина').AsDouble() for el in lines]))

els = FilteredElementCollector(doc)\
    .OfCategory(BuiltInCategory.OST_GenericModel)\
    .WhereElementIsNotElementType().ToElements()

decks = [el for el in els if el.Name == 'Настил']

# levels = {}
# for el in decks:
#     h = round(el.Location.Point.Z * k / 10) * 10
#     if h not in levels:
#         levels[h] = []
#     levels[h].append(el)

lines = FilteredElementCollector(doc)\
    .OfCategory(BuiltInCategory.OST_Lines)\
    .WhereElementIsNotElementType().ToElements()

# levels = {}
# for line in lines:
#     h = round(((line.Location.Curve.GetEndPoint(0) + line.Location.Curve.GetEndPoint(0)) / 2).Z * k / 10) * 10
#     if h not in levels:
#         levels[h] = []
#     levels[h].append(line)

d = 1 / k


def belong(line, deck):
    mx = deck.get_BoundingBox(doc.ActiveView).Max
    mn = deck.get_BoundingBox(doc.ActiveView).Min
    point = (line.Location.Curve.GetEndPoint(0) + line.Location.Curve.GetEndPoint(1)) / 2
    if mn.X + d < point.X < mx.X - d and mn.Y + d < point.Y < mx.Y - d and mn.Z < point.Z < mx.Z:
        return True
    return False


dct = {}
for line in lines:
    if 'Белый' not in line.LineStyle.Name:
        for deck in decks:
            if belong(line, deck):
                if deck.Id not in dct:
                    dct[deck.Id] = []
                dct[deck.Id].append(line)
                break

for deck in decks:
    len = 0
    if deck.Id in dct.keys():
        len = sum([line.LookupParameter('Длина').AsDouble() for line in dct[deck.Id]])
    deck.LookupParameter('Длина линий').Set(len)

# for deckId in dct.keys():
#     deck = doc.GetElement(deckId)
#     deck.LookupParameter('Длина линий').Set(sum([line.LookupParameter('Длина').AsDouble() for line in dct[deckId]]))
    # for line in dct[deckId]:
    #     print('- ', line.Id)

for deck in decks:
    len = 0
    if deck.LookupParameter('в1').AsInteger():
        len += deck.LookupParameter('a1').AsDouble()
        len += deck.LookupParameter('b1').AsDouble()
        len += deck.LookupParameter('Зазор').AsDouble() * 2
    if deck.LookupParameter('в2').AsInteger():
        len += deck.LookupParameter('a2').AsDouble()
        len += deck.LookupParameter('b2').AsDouble()
        len += deck.LookupParameter('Зазор').AsDouble() * 2
    if deck.LookupParameter('в3').AsInteger():
        len += deck.LookupParameter('a3').AsDouble()
        len += deck.LookupParameter('b3').AsDouble()
        len += deck.LookupParameter('Зазор').AsDouble() * 2
    if deck.LookupParameter('в4').AsInteger():
        len += deck.LookupParameter('a4').AsDouble()
        len += deck.LookupParameter('b4').AsDouble()
        len += deck.LookupParameter('Зазор').AsDouble() * 2
    len += deck.LookupParameter('Длина линий').AsDouble()
    deck.LookupParameter('ХТ Длина ОВ').Set(len * k)
    a = deck.LookupParameter('A').AsDouble() * k
    b = deck.LookupParameter('B').AsDouble() * k
    deck.LookupParameter('ХТ Размер фитинга ОВ').Set('{:.0f}x{:.0f}'.format(a, b))
    linesLen = deck.LookupParameter('Длина линий').AsDouble() * k
    s = ''
    if len and linesLen:
        s = '{:.0f} ({:.0f})'.format(len * k, linesLen)
    elif len:
        s = '{:.0f}'.format(len * k)
    elif linesLen:
        s = '({:.0f})'.format(linesLen)
    deck.LookupParameter('Комментарии').Set(s)

t.Commit()
