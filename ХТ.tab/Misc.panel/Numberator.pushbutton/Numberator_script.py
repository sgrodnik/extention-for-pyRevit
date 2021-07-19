# -*- coding: utf-8 -*-
"""Нумератор пространств"""
__title__ = 'Нумератор'
__author__ = 'SG'
import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId
import sys
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# tags = [doc.GetElement(id) for id in uidoc.Selection.GetElementIds() if isinstance(doc.GetElement(id), IndependentTag)]

# if not tags:
#     sys.exit()


class CustomISelectionFilter(ISelectionFilter):
    def __init__(self, nom_categorie):
        self.nom_categorie = nom_categorie

    def AllowElement(self, e):
        if self.nom_categorie in e.Category.Name:
            return True
        else:
            return False

    def AllowReference(self, ref, point):
        return true


# i = 1
# try:
#     target = uidoc.Selection.PickObject(ObjectType.Element, CustomISelectionFilter("Пространства"), 'Выберите пространство')
#     target = doc.GetElement(target.ElementId)
#     i = int(target.LookupParameter('Номер').AsString()) + 1
# except:  # Exceptions.OperationCanceledException:
#     pass

# t = Transaction(doc, 'Перенести марки')
# t.Start()

# while 1:
#     try:
#         target = uidoc.Selection.PickObject(ObjectType.Element, CustomISelectionFilter("Пространства"), 'Выберите пространство')
#     except:  # Exceptions.OperationCanceledException:
#         # t.Commit()
#         break
#         # sys.exit()

#     target = doc.GetElement(target.ElementId)
#     target.LookupParameter('Номер').Set(str(i))
#     i += 1

t = Transaction(doc, 'Нумератор')
t.Start()

sel = [doc.GetElement(elid) for elid in uidoc.Selection.GetElementIds()]

def number_els(base_el, upper=True):
    els = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericAnnotation).OwnedByView(doc.ActiveView.Id).WhereElementIsNotElementType().ToElements()
    els = [i for i in els
            if i.LookupParameter('Номер')
            and i.LookupParameter('Номер').AsString()
            and i.LookupParameter('Номер').AsString().isdigit()
            and i.LookupParameter('Тип').AsValueString() == base_el.LookupParameter('Тип').AsValueString()
          ]
    els = [i for i in els if int(i.LookupParameter('Номер').AsString()) >= int(base_el.LookupParameter('Номер').AsString())]
    for el in sorted(els, key=lambda x: int(x.LookupParameter('Номер').AsString())):
        old_num = int(el.LookupParameter('Номер').AsString())
        new_num = old_num + (1 if upper else -1)
        el.LookupParameter('Номер').Set(str(new_num))
        print(old_num, new_num)
    uidoc.Selection.SetElementIds(List[ElementId]([el.Id for el in els]))



if len(sel) == 1:
    current_num = int(sel[0].LookupParameter('Номер').AsString())
    category_name = sel[0].LookupParameter('Категория').AsValueString()

    if __shiftclick__:
        number_els(sel[0], upper=True)
        t.Commit()
    elif __forceddebugmode__:
        number_els(sel[0], upper=False)
        t.Commit()
    else:
        while True:
            current_num += 1
            try:
                target = uidoc.Selection.PickObject(ObjectType.Element, CustomISelectionFilter(category_name), 'Выберите следующий элемент ' + category_name + ' для ввода "' + str(current_num) + '" [ESC для выхода]')
            except:  # Exceptions.OperationCanceledException:
                t.Commit()
                break
                # sys.exit()

            target = doc.GetElement(target.ElementId)
            target.LookupParameter('Номер').Set(str(current_num))
            doc.Regenerate()
else:
    i = 14
    for el in sel:
        if el.LookupParameter('Номер'):
            el.LookupParameter('Номер').Set(str(i))
        else:
            el.Text = str(i)
        i += 1


    t.Commit()
