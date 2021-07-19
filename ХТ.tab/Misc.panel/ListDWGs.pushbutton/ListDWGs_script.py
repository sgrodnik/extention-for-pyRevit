# -*- coding: UTF-8 -*-
"""Lists all linked and imported DWG instances with worksets and creator."""
import clr
from collections import defaultdict

from pyrevit import revit, DB
from pyrevit import script
from pyrevit import forms


output = script.get_output()


def listdwgs(current_view_only=False):
    dwgs = DB.FilteredElementCollector(revit.doc)\
             .OfClass(DB.ImportInstance)\
             .WhereElementIsNotElementType()\
             .ToElements()

    dwgInst = defaultdict(list)
    workset_table = revit.doc.GetWorksetTable()

    output.print_md("## LINKED AND IMPORTED DWG FILES:")
    output.print_md('By: [{}]({})'.format('Frederic Beaupere',
                                          'https://github.com/frederic-beaupere'))

    for dwg in dwgs:
        if dwg.IsLinked:
            dwgInst["LINKED DWGs:"].append(dwg)
        else:
            dwgInst["IMPORTED DWGs:"].append(dwg)

    for link_mode in dwgInst:
        output.print_md("####{}".format(link_mode))
        for dwg in dwgInst[link_mode]:
            dwg_id = dwg.Id
            dwg_name = \
                dwg.Parameter[DB.BuiltInParameter.IMPORT_SYMBOL_NAME].AsString()
            dwg_workset = workset_table.GetWorkset(dwg.WorksetId).Name
            dwg_instance_creator = \
                DB.WorksharingUtils.GetWorksharingTooltipInfo(revit.doc,
                                                              dwg.Id).Creator

            if current_view_only \
                    and revit.active_view.Id != dwg.OwnerViewId:
                continue

            print('\n\n')
            output.print_md("**DWG name:** {}\n\n"
                            "- DWG created by:{}\n\n"
                            "- DWG id: {}\n\n"
                            "- DWG workset: {}\n\n"
                            .format(dwg_name,
                                    dwg_instance_creator,
                                    output.linkify(dwg_id),
                                    dwg_workset))
    print(output.linkify([el.Id for el in dwgs], 'Select all'))

selected_option = \
    forms.CommandSwitchWindow.show(
        ['In Current View',
         'In Model'],
        message='Select search option:'
        )

if selected_option:
    listdwgs(current_view_only=selected_option == 'In Current View')
