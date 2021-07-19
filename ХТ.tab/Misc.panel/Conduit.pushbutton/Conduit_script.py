# -*- coding: utf-8 -*-
""""""
from __future__ import print_function
__title__ = 'Кабель\nканал'
__author__ = 'SG'

import clr
import re
import os
import sys
import math

clr.AddReference('System.Core')
from System.Collections.Generic import *
from Autodesk.Revit.DB import SectionType, ElementId, PartUtils, ViewOrientation3D, XYZ, FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, Line, Structure, ParameterFilterElement, ParameterFilterRuleFactory, ElementParameterFilter, ViewDuplicateOption
from Autodesk.Revit.DB import WorksharingUtils, Structure
import Autodesk.Revit.DB as DB
import Autodesk.Revit.Exceptions as Exceptions
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from pyrevit import script
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application
output = script.get_output()
k = 304.8


def natural_sorted(list, key=lambda s: s):
    """
    Sort the list into natural alphanumeric order.
    """
    def get_alphanum_key_func(key):
        convert = lambda text: int(text) if text.isdigit() else text  # noqa
        return lambda s: [convert(c) for c in re.split('([0-9]+)', key(s))]
    sort_key = get_alphanum_key_func(key)
    return sorted(list, key=sort_key)


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


sel = [doc.GetElement(elid) for elid in uidoc.Selection.GetElementIds()]

t = Transaction(doc, 'Test')
t.Start()

clr.AddReference('System.Windows.Forms')
clr.AddReference('IronPython.Wpf')
from pyrevit import script

xamlfile = script.get_bundle_file('window.xaml')
import wpf
from System import Windows

# settings = {'var': 3}

import json

param = doc.ProjectInformation.LookupParameter('pyrevit').AsString()
settings = json.loads(param or '{}')

var = settings.get('var', 1)

if __shiftclick__:
    class MyWindow(Windows.Window):
        def __init__(self):
            wpf.LoadComponent(self, xamlfile)
        def var1(self, sender, args):
            self.Close()
            global var
            var = 1
        def var2(self, sender, args):
            self.Close()
            global var
            var = 2
        def var3(self, sender, args):
            self.Close()
            global var
            var = 3
    MyWindow().ShowDialog()

settings['var'] = var
settings = json.dumps(settings, sort_keys=True, indent=4, separators=(',', ': '))
doc.ProjectInformation.LookupParameter('pyrevit').Set(settings)

if var == 1:
    t.SetName('Распространение комментариев')
    comms = []
    engineering_flags = []
    for el in sel:
        engineering_flag = el.LookupParameter('Инженерка').AsInteger() if el.LookupParameter('Инженерка') else 0
        engineering_flags.append(engineering_flag) if engineering_flag == 1 else None
        comment = el.LookupParameter('Комментарии').AsString()
        comment = comment if comment else ''
        comms.append(comment) if comment else None
    comms = list(set(comms))
    engineering_flags = list(set(engineering_flags))
    if len(engineering_flags) == 1:
        for el in sel:
            el.LookupParameter('Инженерка').Set(engineering_flags[0]) if el.LookupParameter('Инженерка') else None
    # else:
    #     print(len(engineering_flags))
    #     print(engineering_flags)
    #     print('--------------')
    if len(comms) == 1:
        for el in sel:
            el.LookupParameter('Комментарии').Set(comms[0])
    else:
        print(len(comms))
        print(comms)
    t.Commit()


elif var == 2:
    t.SetName('Радиусы гофры')
    if len(sel) == 1:
        category_name = sel[0].LookupParameter('Категория').AsValueString()
        radius = sel[0].LookupParameter('Радиус загиба').AsDouble()
        while True:
            radius += 26 / k
            try:
                target = uidoc.Selection.PickObject(ObjectType.Element, CustomISelectionFilter(category_name), 'Выберите следующий элемент ' + category_name + '" [ESC для выхода]')
            except Exceptions.OperationCanceledException:
                t.Commit()
                break
                # sys.exit()

            target = doc.GetElement(target.ElementId)
            target.LookupParameter('Радиус загиба').Set(radius)
            doc.Regenerate()

elif var == 3:
    from collections import Counter
    t.SetName('Сбор комментариев')
    lst = []
    for el in sel:
        comment = el.LookupParameter('Комментарии').AsString()
        comment = comment or ''
        for sub in comment.split(', '):
            if '×' in sub:
                sub, count = sub.split('×')
                for n in range(int(count)):
                    lst.append(sub)
            else:
                # sub, count = [sub, 1]
                lst.append(sub)
            # print(sub + '---' + str(count))
    # for el in Counter(lst).most_common():
    res = ', '.join(['×'.join([el[0], str(el[1])] if el[1] > 1 else [el[0]]) for el in Counter(lst).most_common()])
    print(res)
    # import subprocess
    # subprocess.Popen(['powershell.exe', 'Set-Clipboard -Value "' + res + '"'])
    # import pyperclip
    # pyperclip.copy(res)
    t.Commit()
