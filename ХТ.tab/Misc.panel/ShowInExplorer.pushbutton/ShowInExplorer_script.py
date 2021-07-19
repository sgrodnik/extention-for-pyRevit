# -*- coding: utf-8 -*-
"""Показать файл в проводнике"""
__title__ = 'Показать в\nпроводнике'
__author__ = 'SG'

doc = __revit__.ActiveUIDocument.Document

import shlex, subprocess
args = shlex.split('explorer.exe /select, "{}"'.format(doc.PathName))
subprocess.Popen(args)
