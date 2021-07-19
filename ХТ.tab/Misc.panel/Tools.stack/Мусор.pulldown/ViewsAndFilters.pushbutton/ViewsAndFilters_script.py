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
from Autodesk.Revit.DB import LogicalOrFilter, LogicalAndFilter, OverrideGraphicSettings
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from pyrevit import script
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application
k = 304.8


path = os.path.join(os.path.dirname(doc.PathName), 'ViewsAndFilters.txt')
file_started = False
try:
    src = open(path, 'r').read().decode("utf-8")
    if src:
        src = src[:-1] if src[-1] == '\n' else src
    if not src or src.startswith('\nЭтот'):
        raise IOError
except IOError:
    info = '\nЭтот файл должен содержать таблицу (разделители - табуляция).\nПервая строка должна содержать заголовки. Последняя строка может быть пустой.\n' + path
    with open(path, 'w') as file:
        file.write(info.encode("utf-8"))
    os.startfile(path)
    src = '\n\n'
    file_started = True

if __shiftclick__:
    if not file_started:
        os.startfile(path)
    sys.exit()

src_filters, src_views = src.split('\n\n')
src_filters = src_filters.split('\n')[1:]
src_views = src_views.split('\n')[1:]
if src_filters:
    src_filters.pop(-1) if src_filters[-1].split('\t')[1] == '---' else None
    src_f_ids =   [i.split('\t')[0] for i in src_filters]
    src_f_names = [i.split('\t')[1] for i in src_filters]
    src_f_cats =  [i.split('\t')[2] for i in src_filters]
    src_f_logic = [i.split('\t')[3] for i in src_filters]
    src_f_rules = [i.split('\t')[4] for i in src_filters]
if src_views:
    src_views.pop(-1) if src_views[-1].split('\t')[1] == '---' else None
    src_v_view_id =                  [i.split('\t')[0] for i in src_views]
    src_v_view_symbol =              [i.split('\t')[1] for i in src_views]
    src_v_view_name =                [i.split('\t')[2] for i in src_views]
    src_v_filter_id =                [i.split('\t')[3] for i in src_views]
    src_v_filter_name =              [i.split('\t')[4] for i in src_views]
    src_v_visibility =               [i.split('\t')[5] for i in src_views]
    src_v_pro_line_weight =          [i.split('\t')[6] for i in src_views]
    src_v_pro_line_color =           [i.split('\t')[7] for i in src_views]
    src_v_pro_line_pattern_id =      [i.split('\t')[8] for i in src_views]
    src_v_pro_fill_pattern_visible = [i.split('\t')[9] for i in src_views]
    src_v_pro_fill_pattern_id =      [i.split('\t')[10] for i in src_views]
    src_v_pro_fill_color =           [i.split('\t')[11] for i in src_views]
    src_v_transparency =             [i.split('\t')[12] for i in src_views]
    src_v_cut_line_weight =          [i.split('\t')[13] for i in src_views]
    src_v_cut_line_color =           [i.split('\t')[14] for i in src_views]
    src_v_cut_line_pattern_id =      [i.split('\t')[15] for i in src_views]
    src_v_cut_fill_pattern_visible = [i.split('\t')[16] for i in src_views]
    src_v_cut_fill_pattern_id =      [i.split('\t')[17] for i in src_views]
    src_v_cut_fill_color =           [i.split('\t')[18] for i in src_views]
    src_v_halftone =                 [i.split('\t')[19] for i in src_views]

# data = []
# sheets = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets).WhereElementIsNotElementType().ToElements()
# for sheet in sheets:
#     sheet_number = sheet.SheetNumber
#     sheet_name = sheet.Name
#     viewports = [doc.GetElement(i) for i in sheet.GetAllViewports()]
#     for viewport in viewports:
#         symbol = viewport.Name
#         view_name = doc.GetElement(viewport.ViewId).Name
#         data.append([sheet_number, sheet_name, symbol, view_name, ])

# columns = ['№', 'Имя листа', 'Тип', 'Имя вида', 'Статус']

# output = script.get_output()
# output.print_table(table_data=data, columns=columns)

# sys.exit()


def natural_sorted(list, key=lambda s: s):
    """
    Sort the list into natural alphanumeric order.
    """
    def get_alphanum_key_func(key):
        convert = lambda text: int(text) if text.isdigit() else text  # noqa
        return lambda s: [convert(c) for c in re.split('([0-9]+)', key(s))]
    sort_key = get_alphanum_key_func(key)
    return sorted(list, key=sort_key)


def map_names_to_keys(dict_):
    names = [dict_[i][1] for i in dict_]
    twins = [i for i in names if names.count(i) > 1]
    result = {}
    for key in dict_:
        item, name = dict_[key]
        new_name = (name + ' [' + str(key) + ']') if name in twins else name
        dict_[key] = (item, new_name)
        result[new_name] = item
    return result


# def map_names_to_keys0(dict_):  ## fun. Del.
#     return {(dict_[key][1] + ' [' + str(key) + ']') if list(dict_).count(dict_[key][1]) > 1 else dict_[key][1]: dict_[key][0] for key in dict_}


cats = {}
for cat in doc.Settings.Categories:
    cats[cat.Id.IntegerValue] = (cat, cat.Name)
    if cat.SubCategories:
        for subcat in cat.SubCategories:
            cats[subcat.Id.IntegerValue] = (subcat, subcat.Name)
cats_by_name = map_names_to_keys(cats)

builtin_params = {}
for i in dir(BuiltInParameter):
    try:
        bip = getattr(BuiltInParameter, i)
        name = LabelUtils.GetLabelFor(bip)
        builtin_params[ElementId(bip).IntegerValue] = (bip, name)
    except:
        pass
builtin_params_by_name = map_names_to_keys(builtin_params)

project_params = {}
iterator = doc.ParameterBindings.ForwardIterator()
iterator.Reset()
while(iterator.MoveNext()):
    param = iterator.Key
    project_params[param.Id.IntegerValue] = (param, param.Name)
project_params_by_name = map_names_to_keys(project_params)


def get_param_name(param_id):
    if param_id.IntegerValue < 0:
        param_name = builtin_params[param_id.IntegerValue][1]
    else:
        param_name = project_params[param_id.IntegerValue][1]
    return param_name


evaluator_signs = {'FilterNumericEquals'         : 'n=',  # равно            (число)   # noqa
                   'FilterNumericEquals-'        : 'n!=', # не равно         (число)   # noqa
                   'FilterNumericGreater'        : 'n>',  # больше           (число)   # noqa
                   'FilterNumericGreaterOrEqual' : 'n>=', # больше или равно (число)   # noqa
                   'FilterNumericLess'           : 'n<',  # меньше           (число)   # noqa
                   'FilterNumericLessOrEqual'    : 'n<=', # меньше или равно (число)   # noqa
                   'FilterStringEquals'          : '=',   # равно                      # noqa
                   'FilterStringEquals-'         : '!=',  # не равно                   # noqa
                   'FilterStringGreater'         : '>',   # больше                     # noqa
                   'FilterStringGreaterOrEqual'  : '>=',  # больше или равно           # noqa
                   'FilterStringLess'            : '<',   # меньше                     # noqa
                   'FilterStringLessOrEqual'     : '<=',  # меньше или равно           # noqa
                   'FilterStringContains'        : '^',   # содержит                   # noqa
                   'FilterStringContains-'       : '!^',  # не содержит                # noqa
                   'FilterStringBeginsWith'      : '<<',  # начинается с               # noqa
                   'FilterStringBeginsWith-'     : '!<',  # не начинается с            # noqa
                   'FilterStringEndsWith'        : '>>',  # заканчивается на           # noqa
                   'FilterStringEndsWith-'       : '!>',  # не заканчивается на        # noqa
                  }
pfrf = ParameterFilterRuleFactory
evaluator_methods = {'n=' : pfrf.CreateEqualsRule         ,# равно                      # noqa
                     'n!=': pfrf.CreateNotEqualsRule      ,# не равно                   # noqa
                     'n>' : pfrf.CreateGreaterRule        ,# больше                     # noqa
                     'n>=': pfrf.CreateGreaterOrEqualRule ,# больше или равно           # noqa
                     'n<' : pfrf.CreateLessRule           ,# меньше                     # noqa
                     'n<=': pfrf.CreateLessOrEqualRule    ,# меньше или равно           # noqa
                     '=' : pfrf.CreateEqualsRule          ,# равно                      # noqa
                     '!=': pfrf.CreateNotEqualsRule       ,# не равно                   # noqa
                     '>' : pfrf.CreateGreaterRule         ,# больше                     # noqa
                     '>=': pfrf.CreateGreaterOrEqualRule  ,# больше или равно           # noqa
                     '<' : pfrf.CreateLessRule            ,# меньше                     # noqa
                     '<=': pfrf.CreateLessOrEqualRule     ,# меньше или равно           # noqa
                     '^' : pfrf.CreateContainsRule        ,# содержит                   # noqa
                     '!^': pfrf.CreateNotContainsRule     ,# не содержит                # noqa
                     '<<': pfrf.CreateBeginsWithRule      ,# начинается с               # noqa
                     '!<': pfrf.CreateNotBeginsWithRule   ,# не начинается с            # noqa
                     '>>': pfrf.CreateEndsWithRule        ,# заканчивается на           # noqa
                     '!>': pfrf.CreateNotEndsWithRule     ,# не заканчивается на        # noqa
                     }

# ########################################################################################################################################################
# ########################################################################################################################################################
# ########################################################################################################################################################
# ########################################################################################################################################################
# ########################################################################################################################################################

def get_cats(s):
    cat_ids = []
    for name in s.split('; '):
        cat = cats_by_name.get(name, None)
        if not cat:
            raise Exception('Category "{}" is not found.'.format(name))
        cat_ids.append(cat.Id)
    return cat_ids


def get_param(name):
    param = project_params_by_name.get(name, None) or builtin_params_by_name.get(name, None)
    if not param:
        raise Exception('Parameter "{}" is not found.'.format(name))
    return param


def create_rule(param, evaluator, value):
    param_id = ElementId(param) if isinstance(param, BuiltInParameter) else param.Id
    args = [param_id, value]
    if 'n' in evaluator:
        if '[id' in value:
            value = value.split('[id')[-1]
            value = int(value.replace('[id', '').replace(']', ''))
            args[1] = ElementId(value)
        elif value.isdigit():
            args[1] = int(value)
        else:
            args[1] = float(value)
            args.append(0.00001)
    else:
        args.append(True)
    rule = evaluator_methods[evaluator](*args)
    return rule


def get_rules(s):
    result = []
    for part in s.split('; '):
        param_name = part.split('"')[1]
        value = part.split('"')[3]
        evaluator = part.replace(' "' + param_name + '"', '').replace(' "' + value + '"', '').split()[1]
        param = get_param(param_name)
        rule = create_rule(param, evaluator, value)
        result.append(rule)
    return result


t = Transaction(doc, 'Виды и фильтры')
t.Start()

all_filters = FilteredElementCollector(doc).OfClass(ParameterFilterElement).ToElements()
all_filters = natural_sorted(all_filters, key=lambda x: x.Name)
data = []
for filt in all_filters:
    current_id = str(filt.Id)
    current_name = str(filt.Name)
    current_cats = '; '.join(sorted([cats[elid.IntegerValue][1] for elid in filt.GetCategories()]))
    lfilter = filt.GetElementFilter()
    current_logic = ''
    current_rules = ''
    if lfilter:
        i = 1
        current_logic = lfilter.__class__.__name__.replace('Logical', '').replace('Filter', '')
        cell = []
        # print(current_name)
        if type(lfilter) == ElementParameterFilter:
            continue
        for epfilter in lfilter.GetFilters():  # Сделать для варианта, когда это не лоджикФильтр, а елементФильтр (как в Медси). Или просто пропускать это говно в try?
            try:
                rules = epfilter.GetRules()
            except AttributeError:
                rules = []
                cell = ['Nested filters are not supported']
            if rules:
                for rule in epfilter.GetRules():
                    rulename = rule.__class__.__name__
                    invert = ''
                    if rulename == 'FilterInverseRule':
                        rule = rule.GetInnerRule()
                        invert = '-'
                    param_id = rule.GetRuleParameter()
                    param_name = get_param_name(param_id)
                    evaluator = evaluator_signs.get(str(rule.GetEvaluator().__class__.__name__) + invert, 'Error')
                    try:
                        value = '{:.5f}'.format(rule.RuleValue)
                    except AttributeError:
                        value = str(rule.RuleString)
                    except ValueError:
                        rule_value = rule.RuleValue
                        value = str(rule_value)
                        if type(rule_value) == ElementId:
                            elem = doc.GetElement(rule_value)
                            elem_name = elem.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
                            value = elem_name + ' [id' + str(rule_value) + ']'
                    string = '{}: "{}" {} "{}"'.format(i, param_name, evaluator, value)
                    i += 1
                    cell.append(string)
        current_rules = '; '.join(cell)

    status = []
    if src_filters:
        line_no = src_f_ids.index(current_id) if current_id in src_f_ids else None
        if line_no is not None:
            new_name = src_f_names[line_no]
            name_mod = True if new_name != current_name else False
            current_name = new_name if new_name != current_name else current_name
            new_cats = src_f_cats[line_no]
            cats_mod = True if new_cats != current_cats else False
            current_cats = new_cats if new_cats != current_cats else current_cats
            new_logic = src_f_logic[line_no]
            logic_mod = True if new_logic != current_logic else False
            current_logic = new_logic if new_logic != current_logic else current_logic
            new_rules = src_f_rules[line_no]
            rules_mod = True if new_rules != current_rules else False
            current_rules = new_rules if new_rules != current_rules else current_rules
            if name_mod:
                filt.Name = current_name
                status.append('n')
            if cats_mod:
                filt.SetCategories(List[ElementId](get_cats(current_cats)))
                status.append('c')
            if rules_mod or logic_mod:
                rules = get_rules(current_rules) if current_rules else []
                epfilters = [ElementParameterFilter(rule) for rule in rules]  # не протестирован и не проработан вариант создания фильтра без правил
                lfilter = LogicalAndFilter(epfilters) if current_logic == 'And' else LogicalOrFilter(epfilters)
                filt.SetElementFilter(lfilter)
                status.append('r')
    status = 'Mod ' + ', '.join(status) if status else ''

    data.append([current_id, current_name, current_cats, current_logic, current_rules, status])

if src_filters:
    all_filter_ids = [str(i.Id) for i in all_filters]
    for filter_id in src_f_ids:
        if filter_id not in all_filter_ids:
            line_no = src_f_ids.index(filter_id)
            filter_name = src_f_names[line_no]
            filter_cats = src_f_cats[line_no]
            filter_cats = List[ElementId](get_cats(filter_cats))
            filter_logic = src_f_logic[line_no]
            filter_rules = src_f_rules[line_no]
            filter_rules = get_rules(filter_rules) if filter_rules else []
            epfilters = [ElementParameterFilter(rule) for rule in filter_rules]  # не протестирован и не проработан вариант создания фильтра без правил
            lfilter = LogicalAndFilter(epfilters) if filter_logic == 'And' else LogicalOrFilter(epfilters)
            try:
                pfilter_id = str(ParameterFilterElement.Create(doc, filter_name, filter_cats, lfilter).Id)
                status = 'New'
            except Exception as exception:
                pfilter_id = 'Error'
                src_f_cats[line_no] = exception.message.replace('\n', ' ').replace('\r', ' ')
                status = 'Error'

            data.append([pfilter_id, filter_name, src_f_cats[line_no], filter_logic, src_f_rules[line_no], status])

data.append(['---', '---', '---', '---', '---', len([i for i in data if i[5] != ''])])

columns = ['Id фильтра', 'Имя фильтра', 'Категории', 'Логика', 'Правила', 'Статус']

output = script.get_output()
output.print_table(table_data=data, columns=columns)

# ########################################################################################################################################################
# ########################################################################################################################################################
# ########################################################################################################################################################
# ########################################################################################################################################################
# ########################################################################################################################################################

def col(s):
    arr = [int(i) for i in s.split()]
    return Color(*arr) if len(arr) == 3 else None


def line(s):
    if s == '-1':
        return None
    all_patterns = FilteredElementCollector(doc).OfClass(LinePatternElement).ToElements()
    return [i for i in all_patterns if i.Name == s][0].Id


def fill(s):
    if s == '-1':
        return None
    all_patterns = FilteredElementCollector(doc).OfClass(FillPatternElement).ToElements()
    return [i for i in all_patterns if i.Name == s][0].Id

done = []
views = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views).WhereElementIsNotElementType().ToElements()
views = [i for i in views if i.Origin]
views = natural_sorted(views, key=lambda x: x.Name)
views = sorted(views, key=lambda x: x.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString())
data = []
for view in views:
    view_filters = [doc.GetElement(i) for i in view.GetFilters()]
    if view_filters:
        for view_filter in view_filters:
            cfg = view.GetFilterOverrides(view_filter.Id)
            current_view_id = str(view.Id)
            current_view_symbol = view.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
            current_view_name = view.Name
            current_filter_id = str(view_filter.Id)
            current_filter_name = str(view_filter.Name)
            current_visibility = str(1 if view.GetFilterVisibility(view_filter.Id) else 0)
            current_pro_line_weight = str(cfg.ProjectionLineWeight)
            current_pro_line_color = str('{} {} {}'.format(cfg.ProjectionLineColor.Red, cfg.ProjectionLineColor.Green, cfg.ProjectionLineColor.Blue) if cfg.ProjectionLineColor.IsValid else '-1')
            current_pro_line_pattern_id = str(doc.GetElement(cfg.ProjectionLinePatternId).Name if cfg.ProjectionLinePatternId.IntegerValue > 0 else '-1')
            current_pro_fill_pattern_visible = str(1 if cfg.IsProjectionFillPatternVisible else 0)
            current_pro_fill_pattern_id = str(doc.GetElement(cfg.ProjectionFillPatternId).Name if cfg.ProjectionFillPatternId.IntegerValue > 0 else '-1')
            current_pro_fill_color = str('{} {} {}'.format(cfg.ProjectionFillColor.Red, cfg.ProjectionFillColor.Green, cfg.ProjectionFillColor.Blue) if cfg.ProjectionFillColor.IsValid else '-1')
            current_transparency = str(cfg.Transparency)
            current_cut_line_weight = str(cfg.CutLineWeight)
            current_cut_line_color = str('{} {} {}'.format(cfg.CutLineColor.Red, cfg.CutLineColor.Green, cfg.CutLineColor.Blue) if cfg.CutLineColor.IsValid else '-1')
            current_cut_line_pattern_id = str(doc.GetElement(cfg.CutLinePatternId).Name if cfg.CutLinePatternId.IntegerValue > 0 else '-1')
            current_cut_fill_pattern_visible = str(1 if cfg.IsCutFillPatternVisible else 0)
            current_cut_fill_pattern_id = str(doc.GetElement(cfg.CutFillPatternId).Name if cfg.CutFillPatternId.IntegerValue > 0 else '-1')
            current_cut_fill_color = str('{} {} {}'.format(cfg.CutFillColor.Red, cfg.CutFillColor.Green, cfg.CutFillColor.Blue) if cfg.CutFillColor.IsValid else '-1')
            current_halftone = str(1 if cfg.Halftone else 0)
            done.append((current_view_id, current_filter_id))

            status = ''
            if src_views:
                line_no = None
                for row in src_views:
                    if current_view_id in row and current_filter_id in row:
                        line_no = src_views.index(row)
                if line_no is not None:
                    new_view_symbol = src_v_view_symbol[line_no]
                    view_symbol_mod = True if new_view_symbol != current_view_symbol else False  ########################################## дописать
                    current_view_symbol = new_view_symbol if new_view_symbol != current_view_symbol else current_view_symbol

                    new_view_name = src_v_view_name[line_no]
                    view_name_mod = True if new_view_name != current_view_name else False
                    current_view_name = new_view_name if new_view_name != current_view_name else current_view_name

                    new_visibility = src_v_visibility[line_no]
                    visibility_mod = True if new_visibility != current_visibility else False
                    current_visibility = new_visibility if new_visibility != current_visibility else current_visibility

                    new_pro_line_weight = src_v_pro_line_weight[line_no]
                    cfg_mod = True if new_pro_line_weight != current_pro_line_weight else False
                    current_pro_line_weight = new_pro_line_weight if new_pro_line_weight != current_pro_line_weight else current_pro_line_weight

                    new_pro_line_color = src_v_pro_line_color[line_no]
                    cfg_mod = True if new_pro_line_color != current_pro_line_color else cfg_mod
                    current_pro_line_color = new_pro_line_color if new_pro_line_color != current_pro_line_color else current_pro_line_color

                    new_pro_line_pattern_id = src_v_pro_line_pattern_id[line_no]
                    cfg_mod = True if new_pro_line_pattern_id != current_pro_line_pattern_id else cfg_mod
                    current_pro_line_pattern_id = new_pro_line_pattern_id if new_pro_line_pattern_id != current_pro_line_pattern_id else current_pro_line_pattern_id

                    new_pro_fill_pattern_visible = src_v_pro_fill_pattern_visible[line_no]
                    cfg_mod = True if new_pro_fill_pattern_visible != current_pro_fill_pattern_visible else cfg_mod
                    current_pro_fill_pattern_visible = new_pro_fill_pattern_visible if new_pro_fill_pattern_visible != current_pro_fill_pattern_visible else current_pro_fill_pattern_visible

                    new_pro_fill_pattern_id = src_v_pro_fill_pattern_id[line_no]
                    cfg_mod = True if new_pro_fill_pattern_id != current_pro_fill_pattern_id else cfg_mod
                    current_pro_fill_pattern_id = new_pro_fill_pattern_id if new_pro_fill_pattern_id != current_pro_fill_pattern_id else current_pro_fill_pattern_id

                    new_pro_fill_color = src_v_pro_fill_color[line_no]
                    cfg_mod = True if new_pro_fill_color != current_pro_fill_color else cfg_mod
                    current_pro_fill_color = new_pro_fill_color if new_pro_fill_color != current_pro_fill_color else current_pro_fill_color

                    new_transparency = src_v_transparency[line_no]
                    cfg_mod = True if new_transparency != current_transparency else cfg_mod
                    current_transparency = new_transparency if new_transparency != current_transparency else current_transparency

                    new_cut_line_weight = src_v_cut_line_weight[line_no]
                    cfg_mod = True if new_cut_line_weight != current_cut_line_weight else cfg_mod
                    current_cut_line_weight = new_cut_line_weight if new_cut_line_weight != current_cut_line_weight else current_cut_line_weight

                    new_cut_line_color = src_v_cut_line_color[line_no]
                    cfg_mod = True if new_cut_line_color != current_cut_line_color else cfg_mod
                    current_cut_line_color = new_cut_line_color if new_cut_line_color != current_cut_line_color else current_cut_line_color

                    new_cut_line_pattern_id = src_v_cut_line_pattern_id[line_no]
                    cfg_mod = True if new_cut_line_pattern_id != current_cut_line_pattern_id else cfg_mod
                    current_cut_line_pattern_id = new_cut_line_pattern_id if new_cut_line_pattern_id != current_cut_line_pattern_id else current_cut_line_pattern_id

                    new_cut_fill_pattern_visible = src_v_cut_fill_pattern_visible[line_no]
                    cfg_mod = True if new_cut_fill_pattern_visible != current_cut_fill_pattern_visible else cfg_mod
                    current_cut_fill_pattern_visible = new_cut_fill_pattern_visible if new_cut_fill_pattern_visible != current_cut_fill_pattern_visible else current_cut_fill_pattern_visible

                    new_cut_fill_pattern_id = src_v_cut_fill_pattern_id[line_no]
                    cfg_mod = True if new_cut_fill_pattern_id != current_cut_fill_pattern_id else cfg_mod
                    current_cut_fill_pattern_id = new_cut_fill_pattern_id if new_cut_fill_pattern_id != current_cut_fill_pattern_id else current_cut_fill_pattern_id

                    new_cut_fill_color = src_v_cut_fill_color[line_no]
                    cfg_mod = True if new_cut_fill_color != current_cut_fill_color else cfg_mod
                    current_cut_fill_color = new_cut_fill_color if new_cut_fill_color != current_cut_fill_color else current_cut_fill_color

                    new_halftone = src_v_halftone[line_no]
                    cfg_mod = True if new_halftone != current_halftone else cfg_mod
                    current_halftone = new_halftone if new_halftone != current_halftone else current_halftone

                    if current_view_symbol == 'del':
                        view.RemoveFilter(ElementId(int(current_filter_id)))
                        status = 'del'
                    else:
                        if view_name_mod:
                            view.Name = current_view_name
                            status = 'Mod'
                        if visibility_mod:
                            view.SetFilterVisibility(ElementId(int(current_filter_id)), bool(int(filter_visibility)))
                            status = 'Mod'
                        if cfg_mod:
                            cfg.SetProjectionLineWeight(              int(current_pro_line_weight)           ) if      int(current_pro_line_weight)     > 0 else None  # noqa
                            cfg.SetProjectionLineColor(               col(current_pro_line_color)            ) if      col(current_pro_line_color)          else None  # noqa
                            cfg.SetProjectionLinePatternId(          line(current_pro_line_pattern_id)       ) if     line(current_pro_line_pattern_id)     else None  # noqa
                            cfg.SetProjectionFillPatternVisible( bool(int(current_pro_fill_pattern_visible)) )                                                         # noqa
                            cfg.SetProjectionFillPatternId(          fill(current_pro_fill_pattern_id)       ) if     fill(current_pro_fill_pattern_id)     else None  # noqa
                            cfg.SetProjectionFillColor(               col(current_pro_fill_color)            ) if      col(current_pro_fill_color)          else None  # noqa
                            cfg.SetSurfaceTransparency(               int(current_transparency)              )                                                         # noqa
                            cfg.SetCutLineWeight(                     int(current_cut_line_weight)           ) if      int(current_cut_line_weight)     > 0 else None  # noqa
                            cfg.SetCutLineColor(                      col(current_cut_line_color)            ) if      col(current_cut_line_color)          else None  # noqa
                            cfg.SetCutLinePatternId(                 line(current_cut_line_pattern_id)       ) if     line(current_cut_line_pattern_id)     else None  # noqa
                            cfg.SetCutFillPatternVisible(        bool(int(current_cut_fill_pattern_visible)) )                                                         # noqa
                            cfg.SetCutFillPatternId(                 fill(current_cut_fill_pattern_id)       ) if     fill(current_cut_fill_pattern_id)     else None  # noqa
                            cfg.SetCutFillColor(                      col(current_cut_fill_color)            ) if      col(current_cut_fill_color)          else None  # noqa
                            cfg.SetHalftone(                     bool(int(current_halftone))                 )                                                         # noqa
                            view.SetFilterOverrides(ElementId(int(current_filter_id)), cfg)
                            status = 'Mod'

            data.append([current_view_id, current_view_symbol, current_view_name, current_filter_id, current_filter_name,
                         current_visibility, current_pro_line_weight, current_pro_line_color, current_pro_line_pattern_id,
                         current_pro_fill_pattern_visible, current_pro_fill_pattern_id, current_pro_fill_color, current_transparency,
                         current_cut_line_weight, current_cut_line_color, current_cut_line_pattern_id, current_cut_fill_pattern_visible,
                         current_cut_fill_pattern_id, current_cut_fill_color, current_halftone, status])
    else:
        current_view_id = str(view.Id)
        current_view_symbol = view.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
        current_view_name = view.Name
        data.append([current_view_id, current_view_symbol, current_view_name, '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''])

all_filters = FilteredElementCollector(doc).OfClass(ParameterFilterElement).ToElements()
if src_views:
    view_ids = [str(i.Id) for i in views]
    for (line_no, (view_id, filter_id)) in enumerate(zip(src_v_view_id, src_v_filter_id)):
        if (view_id, filter_id) not in done:
            # print(view_id, filter_id)
            # line_no = src_v_view_id.index(view_id)

            orig_view_name =           src_v_view_symbol[line_no]
            view_id =                  src_v_view_id[line_no]
            view_name =                src_v_view_name[line_no]
            filter_id =                src_v_filter_id[line_no]
            filter_name =              src_v_filter_name[line_no]
            visibility =               src_v_visibility[line_no]               if filter_name else ''
            pro_line_weight =          src_v_pro_line_weight[line_no]          if filter_name else ''
            pro_line_color =           src_v_pro_line_color[line_no]           if filter_name else ''
            pro_line_pattern_id =      src_v_pro_line_pattern_id[line_no]      if filter_name else ''
            pro_fill_pattern_visible = src_v_pro_fill_pattern_visible[line_no] if filter_name else ''
            pro_fill_pattern_id =      src_v_pro_fill_pattern_id[line_no]      if filter_name else ''
            pro_fill_color =           src_v_pro_fill_color[line_no]           if filter_name else ''
            transparency =             src_v_transparency[line_no]             if filter_name else ''
            cut_line_weight =          src_v_cut_line_weight[line_no]          if filter_name else ''
            cut_line_color =           src_v_cut_line_color[line_no]           if filter_name else ''
            cut_line_pattern_id =      src_v_cut_line_pattern_id[line_no]      if filter_name else ''
            cut_fill_pattern_visible = src_v_cut_fill_pattern_visible[line_no] if filter_name else ''
            cut_fill_pattern_id =      src_v_cut_fill_pattern_id[line_no]      if filter_name else ''
            cut_fill_color =           src_v_cut_fill_color[line_no]           if filter_name else ''
            halftone =                 src_v_halftone[line_no]                 if filter_name else ''

            if orig_view_name:
                print(orig_view_name)
                if view_name in [i.Name for i in views]:
                    # status = 'Name error'
                    # data.append([view_id, orig_view_name, view_name] + ['' for i in range(17)] + [status])
                    continue
                orig_view = [i for i in views if i.Name == orig_view_name][0]
                view = doc.GetElement(orig_view.Duplicate(ViewDuplicateOption.Duplicate))
            else:
                view = [i for i in views if str(i.Id) == view_id][0]

            try:
            	view.Name = view_name
            except Exception as e:
            	print('view_name')
                print(view.Name)
            	print(view_name)
            	raise e
            view_id = view.Id
            view_symbol = view.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()

            if filter_name:
                filter_id = str([i.Id for i in all_filters if i.Name == filter_name][0])
                view.SetFilterVisibility(ElementId(int(filter_id)), bool(int(visibility)))
                cfg = OverrideGraphicSettings()

                cfg.SetProjectionLineWeight(              int(pro_line_weight)           ) if      int(pro_line_weight)     > 0 else None  # noqa
                cfg.SetProjectionLineColor(               col(pro_line_color)            ) if      col(pro_line_color)          else None  # noqa
                cfg.SetProjectionLinePatternId(          line(pro_line_pattern_id)       ) if     line(pro_line_pattern_id)     else None  # noqa
                cfg.SetProjectionFillPatternVisible( bool(int(pro_fill_pattern_visible)) )                                                         # noqa
                cfg.SetProjectionFillPatternId(          fill(pro_fill_pattern_id)       ) if     fill(pro_fill_pattern_id)     else None  # noqa
                cfg.SetProjectionFillColor(               col(pro_fill_color)            ) if      col(pro_fill_color)          else None  # noqa
                cfg.SetSurfaceTransparency(               int(transparency)              )                                                         # noqa
                cfg.SetCutLineWeight(                     int(cut_line_weight)           ) if      int(cut_line_weight)     > 0 else None  # noqa
                cfg.SetCutLineColor(                      col(cut_line_color)            ) if      col(cut_line_color)          else None  # noqa
                cfg.SetCutLinePatternId(                 line(cut_line_pattern_id)       ) if     line(cut_line_pattern_id)     else None  # noqa
                cfg.SetCutFillPatternVisible(        bool(int(cut_fill_pattern_visible)) )                                                         # noqa
                cfg.SetCutFillPatternId(                 fill(cut_fill_pattern_id)       ) if     fill(cut_fill_pattern_id)     else None  # noqa
                cfg.SetCutFillColor(                      col(cut_fill_color)            ) if      col(cut_fill_color)          else None  # noqa
                cfg.SetHalftone(                     bool(int(halftone))                 )                                                         # noqa

                view.SetFilterOverrides(ElementId(int(filter_id)), cfg)
            status = 'New'
            data.append([view_id, view_symbol, view_name, filter_id, filter_name,
                         visibility, pro_line_weight, pro_line_color, pro_line_pattern_id,
                         pro_fill_pattern_visible, pro_fill_pattern_id, pro_fill_color, transparency,
                         cut_line_weight, cut_line_color, cut_line_pattern_id, cut_fill_pattern_visible,
                         cut_fill_pattern_id, cut_fill_color, halftone, status])


#     # doc.ActiveView.Duplicate(name).SetFilterVisibility([i.Id for i in all_filters if i.Name == name + ' не'][0], False)
#     new_view = doc.GetElement(doc.ActiveView.Duplicate(ViewDuplicateOption.Duplicate))
#     new_view.Name = name
#     new_view.SetFilterVisibility(pfilter.Id, False)

n = [len([i for i in data if i[20] != ''])]
data.append(n + ['---' for i in range(19)] + n)

columns = ['Id вида', 'Тип вида', 'Имя вида', 'Id фильтра', 'Имя фильтра', 'Видимость', 'Проекция. Линии. Вес',
           'Проекция. Линии. Цвет', 'Проекция. Линии. Образец', 'Проекция. Штриховки. Видимость', 'Проекция. Штриховки. Образец',
           'Проекция. Штриховки. Цвет', 'Проекция. Прозрачность', 'Разрез. Линии. Вес', 'Разрез. Линии. Цвет', 'Разрез. Линии. Образец',
           'Разрез. Штриховки. Видимость', 'Разрез. Штриховки. Образец', 'Разрез. Штриховки. Цвет', 'Полутона', 'Статус']

if not data:
    data = [columns]
output.print_table(table_data=data, columns=columns)

# ########################################################################################################################################################
# ########################################################################################################################################################
# ########################################################################################################################################################
# ########################################################################################################################################################
# ########################################################################################################################################################

t.Commit()
