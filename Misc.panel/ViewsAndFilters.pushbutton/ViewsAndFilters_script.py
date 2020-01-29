# -*- coding: utf-8 -*-
""""""
__title__ = 'Виды и\nфильтры'
__author__ = 'SG'

import os
import sys
import re
import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *
from Autodesk.Revit.DB import ElementId, FilteredElementCollector, BuiltInCategory, Transaction, BuiltInParameter, ParameterFilterElement, ElementParameterFilter, ViewDuplicateOption
from Autodesk.Revit.DB import Color, LinePatternElement, FillPatternElement, ParameterFilterRuleFactory, ParameterElement, InternalDefinition, LabelUtils, SharedParameterElement
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from pyrevit import script
# from Autodesk.Revit.ApplicationServices.Application import Create
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application
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


cats_by_name = {}
cats = {}
for cat in doc.Settings.Categories:
    # print('{} {}'.format(cat.Id.IntegerValue, cat.Name))
    cats[cat.Id.IntegerValue] = (cat, cat.Name)
    cats_by_name[cat.Name] = cat
    if cat.SubCategories:
        for subcat in cat.SubCategories:
            # print('{} {}'.format(subcat.Id.IntegerValue, subcat.Name))
            cats[subcat.Id.IntegerValue] = (subcat, subcat.Name)
            cats_by_name[cat.Name] = subcat

builtin_params = {}
for i in dir(BuiltInParameter):
    try:
        bip = getattr(BuiltInParameter, i)
        builtin_params[ElementId(bip).IntegerValue] = (bip, LabelUtils.GetLabelFor(bip))
    except:
        pass

project_params = {}
project_param_names = []
iterator = doc.ParameterBindings.ForwardIterator()
iterator.Reset()
while(iterator.MoveNext()):
    project_params[iterator.Key.Id.IntegerValue] = (iterator.Key, iterator.Key.Name)
    project_param_names.append(iterator.Key.Name)

twins = []
for name in project_param_names:
    if project_param_names.count(name) > 1:
        twins.append(name)

for key in project_params:  # Name individualizing
    param = project_params[key][0]
    name = project_params[key][1]
    if name in twins:
        project_params[key] = (param, name + ' (id ' + str(key) + ')')


def get_param_name(param_id):
    if param_id.IntegerValue < 0:
        param_name = builtin_params[param_id.IntegerValue][1]
    else:
        param_name = project_params[param_id.IntegerValue][1]
        # param_name = doc.GetElement(param_id).Name
    return param_name


evaluators = {'FilterNumericEquals'         : '=',   # равно            (число)
              'FilterNumericEquals-'        : '!=',  # не равно         (число)
              'FilterNumericGreater'        : '>',   # больше           (число)
              'FilterNumericGreaterOrEqual' : '>=',  # больше или равно (число)
              'FilterNumericLess'           : '<',   # меньше           (число)
              'FilterNumericLessOrEqual'    : '<=',  # меньше или равно (число)
              'FilterStringEquals'          : '=',   # равно
              'FilterStringEquals-'         : '!=',  # не равно
              'FilterStringGreater'         : '>',   # больше
              'FilterStringGreaterOrEqual'  : '>=',  # больше или равно
              'FilterStringGreater'         : '<',   # меньше
              'FilterStringGreaterOrEqual'  : '<=',  # меньше или равно
              'FilterStringContains'        : '^',   # содержит
              'FilterStringContains-'       : '!^',  # не содержит
              'FilterStringBeginsWith'      : '<<',  # начинается с
              'FilterStringBeginsWith-'     : '!<',  # не начинается с
              'FilterStringEndsWith'        : '>>',  # заканчивается на
              'FilterStringEndsWith-'       : '!>',  # не заканчивается на
              }

# sys.exit()

all_filters = FilteredElementCollector(doc).OfClass(ParameterFilterElement).ToElements()
all_filters = natural_sorted(all_filters, key=lambda x: x.Name)
data = []
for filt in all_filters:
    d = []
    d.append(str(filt.Id))
    d.append(str(filt.Name))
    arrr = [cats[elid.IntegerValue][1] for elid in filt.GetCategories()]
    d.append(str('; '.join(sorted(arrr))))
    # print('filt ' + filt.Name)
    logical_filter = filt.GetElementFilter()
    d.append(str(logical_filter))
    if logical_filter:
        i = 1
        dd = []
        dd.append(logical_filter.__class__.__name__)
        for epfilter in logical_filter.GetFilters():
            for rule in epfilter.GetRules():
                rulename = rule.__class__.__name__
                invert = ''
                if rulename == 'FilterInverseRule':
                    rule = rule.GetInnerRule()
                    # rulename = rule.__class__.__name__ + ' I'
                    invert = '-'
                param_id = rule.GetRuleParameter()
                param_name = get_param_name(param_id)
                evaluator = evaluators.get(str(rule.GetEvaluator().__class__.__name__) + invert, 'Error')
                # dd.append(str(i) + ') ' + str(rule.GetRuleParameter()))
                # dd.append(str(get_param_name(param_id)))
                # dd.append(evaluator)
                try:
                    value = '{:.5f}'.format(rule.RuleValue)
                    # dd.append('{:.5f}'.format(rule.RuleValue))
                except ValueError:
                    value = '{}'.format(rule.RuleValue)
                    # dd.append('{}'.format(rule.RuleValue))
                except AttributeError:
                    value = str(rule.RuleString)
                string = '{}: "{}" {} {}'.format(i, param_name, evaluator, value)
                i += 1
                dd.append(string)
        d.append(str('; '.join(dd)))
    # else:
    #     d.append('-')
    data.append(d)

columns = ['Id фильтра',
           'Имя фильтра',
           'Категории',
           'Правила']

output = script.get_output()
output.print_table(table_data=data, columns=columns)

# sys.exit()

views = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views).WhereElementIsNotElementType().ToElements()
views = [i for i in views if i.Origin]
views = natural_sorted(views, key=lambda x: x.Name)
views = sorted(views, key=lambda x: x.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString())
data = []
for view in views:
    view_filters = [doc.GetElement(i) for i in view.GetFilters()]
    for view_filter in view_filters:
        d = []
        d.append(str(view.Id))
        d.append(view.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString())
        d.append(view.Name)
        d.append(str(view_filter.Id))
        d.append(str(view_filter.Name))
        d.append(str(1 if view.GetFilterVisibility(view_filter.Id) else 0))
        cfg = view.GetFilterOverrides(view_filter.Id)
        d.append(str(cfg.ProjectionLineWeight))
        d.append(str('{} {} {}'.format(cfg.ProjectionLineColor.Red, cfg.ProjectionLineColor.Green, cfg.ProjectionLineColor.Blue) if cfg.ProjectionLineColor.IsValid else '-1'))
        d.append(str(doc.GetElement(cfg.ProjectionLinePatternId).Name if cfg.ProjectionLinePatternId.IntegerValue > 0 else '-1'))
        d.append(str(1 if cfg.IsProjectionFillPatternVisible else 0))
        d.append(str(doc.GetElement(cfg.ProjectionFillPatternId).Name if cfg.ProjectionFillPatternId.IntegerValue > 0 else '-1'))
        d.append(str('{} {} {}'.format(cfg.ProjectionFillColor.Red, cfg.ProjectionFillColor.Green, cfg.ProjectionFillColor.Blue) if cfg.ProjectionFillColor.IsValid else '-1'))
        d.append(str(cfg.Transparency))
        d.append(str(cfg.CutLineWeight))
        d.append(str('{} {} {}'.format(cfg.CutLineColor.Red, cfg.CutLineColor.Green, cfg.CutLineColor.Blue) if cfg.CutLineColor.IsValid else '-1'))
        d.append(str(doc.GetElement(cfg.CutLinePatternId).Name if cfg.CutLinePatternId.IntegerValue > 0 else '-1'))
        d.append(str(1 if cfg.IsCutFillPatternVisible else 0))
        d.append(str(doc.GetElement(cfg.CutFillPatternId).Name if cfg.CutFillPatternId.IntegerValue > 0 else '-1'))
        d.append(str('{} {} {}'.format(cfg.CutFillColor.Red, cfg.CutFillColor.Green, cfg.CutFillColor.Blue) if cfg.CutFillColor.IsValid else '-1'))
        d.append(str(1 if cfg.Halftone else 0))
        data.append(d)

columns = ['Id вида',
           'Тип вида',
           'Имя вида',
           'Id фильтра',
           'Имя фильтра',
           'Видимость',
           'Проекция. Линии. Вес',
           'Проекция. Линии. Цвет',
           'Проекция. Линии. Образец',
           'Проекция. Штриховки. Видимость',
           'Проекция. Штриховки. Образец',
           'Проекция. Штриховки. Цвет',
           'Проекция. Прозрачность',
           'Разрез. Линии. Вес',
           'Разрез. Линии. Цвет',
           'Разрез. Линии. Образец',
           'Разрез. Штриховки. Видимость',
           'Разрез. Штриховки. Образец',
           'Разрез. Штриховки. Цвет',
           'Полутона']

output.print_table(table_data=data, columns=columns)

path = os.path.join(os.path.dirname(doc.PathName), 'ViewsAndFilters.txt')
try:
    src = open(path, 'r').read().decode("utf-8")
    if src:
        src = src[:-1] if src[-1] == '\n' else src
    if not src or src[0] == '\n':
        raise IOError
except IOError:
    info = '\n' + path + '\nЭтот файл должен содержать таблицу (разделители - табуляция).\nПервая строка должна содержать заголовки. Последняя строка может быть пустой.\n'
    with open(path, 'w') as file:
        file.write(info.encode("utf-8"))
    os.startfile(path)
    sys.exit()

if __shiftclick__:
        os.startfile(path)
        sys.exit()

src_filters, src_views = src.split('\n\n')


def get_cats(s):
    return [cats_by_name[name] for name in s.split('; ')]


# def get_rules(s):
#     for part in s.split(': ')[1:]:
#         param_id =
#         param_name =
#         evaluator =
#         value =


for string in src_filters.split('\n')[1:]:
    string = string.split('\t')
    filter_id =                         ElementId(int(string[0]))
    filter_name =                                     string[2]
    cats =                                   get_cats(string[3])
    and_or_or =                                       string[4].split(': ')[0]
    rules =                                 get_rules(string[4])


t = Transaction(doc, 'Виды и фильтры')
t.Start()


def col(s):
    arr = [int(i) for i in s.split()]
    return Color(*arr) if len(arr) == 3 else None


def line(s):
    if s == '-1':
        return None
    all_patterns = FilteredElementCollector(doc).OfClass(LinePatternElement).ToElements()
    # print(s)
    return [i for i in all_patterns if i.Name == s][0].Id


def fill(s):
    if s == '-1':
        return None
    all_patterns = FilteredElementCollector(doc).OfClass(FillPatternElement).ToElements()
    return [i for i in all_patterns if i.Name == s][0].Id


for string in src_views.split('\n')[1:]:
    string = string.split('\t')
    view_id =                                  ElementId(int(string[0]))
    view_symbol =                                            string[1]
    view_name =                                              string[2]
    view_filter_id =                           ElementId(int(string[3]))
    view_filter_name =                                       string[4]
    filter_visibility =                             bool(int(string[5]))
    projectionLineWeight =                               int(string[6])
    projectionLineColor =                                col(string[7])
    projectionLinePatternId =                           line(string[8])
    isProjectionFillPatternVisible =                bool(int(string[9]))
    projectionFillPatternId =                           fill(string[10])
    projectionFillColor =                                col(string[11])
    transparency =                                       int(string[12])
    cutLineWeight =                                      int(string[13])
    cutLineColor =                                       col(string[14])
    cutLinePatternId =                                  line(string[15])
    isCutFillPatternVisible =                       bool(int(string[16]))
    cutFillPatternId =                                  fill(string[17])
    cutFillColor =                                       col(string[18])
    halftone =                                      bool(int(string[19]))


    view = doc.GetElement(view_id)
    view_filter = doc.GetElement(view_filter_id)
    cfg = view.GetFilterOverrides(view_filter.Id)
    # cfg = OverrideGraphicSettings()
    view.SetFilterVisibility(view_filter_id, filter_visibility)
    cfg.SetProjectionLineWeight(projectionLineWeight) if projectionLineWeight > 0 else None
    cfg.SetProjectionLineColor(projectionLineColor) if projectionLineColor else None
    cfg.SetProjectionLinePatternId(projectionLinePatternId) if projectionLinePatternId else None
    cfg.SetProjectionFillPatternVisible(isProjectionFillPatternVisible)
    cfg.SetProjectionFillPatternId(projectionFillPatternId) if projectionFillPatternId else None
    cfg.SetProjectionFillColor(projectionFillColor) if projectionFillColor else None
    cfg.SetSurfaceTransparency(transparency)
    cfg.SetCutLineWeight(cutLineWeight) if cutLineWeight > 0 else None
    cfg.SetCutLineColor(cutLineColor) if cutLineColor else None
    cfg.SetCutLinePatternId(cutLinePatternId) if cutLinePatternId else None
    cfg.SetCutFillPatternVisible(isCutFillPatternVisible)
    cfg.SetCutFillPatternId(cutFillPatternId) if cutFillPatternId else None
    cfg.SetCutFillColor(cutFillColor) if cutFillColor else None
    cfg.SetHalftone(halftone)
    view.SetFilterOverrides(view_filter_id, cfg)

t.Commit()