# -*- coding: utf-8 -*-
""""""
__title__ = 'Копирование\nв сеть'
__author__ = 'SG'
import os
import shutil
import clr
clr.AddReference('System.Core')
from System.Collections.Generic import *
from Autodesk.Revit.DB import Mechanical, FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, BuiltInParameter, ElementId
import sys
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
import codecs
import time
import datetime

from pyrevit import forms


# doc_path_name = doc.PathName.split(' ОВ v')[0] + ' путь в обмен.txt'
doc_pathname = doc.PathName
# print(doc_pathname)
doc_dirname = os.path.dirname(doc_pathname)
# print(doc_dirname)
doc_basename = os.path.basename(doc_pathname)
new_basename = doc_basename.split(' ОВ v')[0] + ' ОВ.rvt'
# print(doc_basename)
# print(os.path.join(doc_dirname, doc_basename))
# print(os.path.dirname(path))

storage_path = os.path.join(doc_dirname, doc_basename.split(' ОВ v')[0] + ' Путь в сети.txt')
try:
    remote_path = open(storage_path, 'r').read().decode("utf-8")
    # print(remote_path)
except IOError:
    remote_path = forms.save_file(
        file_ext='',
        default_name=new_basename,
        init_dir=doc_dirname,
        files_filter='')
    if remote_path:
        with codecs.open(storage_path, mode='w', encoding='utf-8') as f:
            f.write(remote_path)
    else:
        sys.exit()

if os.access(remote_path, os.F_OK):
    old_path = os.path.join(os.path.dirname(remote_path), 'old')
    # print(old_path)
    if not os.access(old_path, os.F_OK):
        os.mkdir(old_path)
    basename = os.path.basename(remote_path)
    timestamp = os.path.getctime(remote_path)
    (mode, ino, dev, nlink, uid, gid, size, atime, mtime, ctime) = os.stat(remote_path)
    modtime = datetime.datetime.fromtimestamp(mtime)
    timestamp = ' (от ' + str(modtime).replace('-', '.').replace(':', '-') + ')'
    # print(timestamp)
    reserve_basename = os.path.splitext(basename)[0] + timestamp + os.path.splitext(basename)[1]
    result = shutil.copy2(remote_path, os.path.join(old_path, reserve_basename))
    print('Файл {}\nскопирован в {}\n---'.format(remote_path, os.path.join(old_path, reserve_basename)))

try:
    shutil.copy2(doc_pathname, remote_path)
    print('Файл {}\nскопирован в {}'.format(doc_pathname, remote_path))
except IOError as error:
    print('Ошибка записи')
    print(error.strerror)


# os.path.splitext(filename)[1]

# els = [doc.GetElement(id) for id in uidoc.Selection.GetElementIds()]

# t = Transaction(doc, 'CopyFile')
# t.Start()

# src = doc.PathName
# path = os.path.dirname(src)
# dst = path + '123.rvt'
# # print(path + '123.rvt')
# for i in os.listdir(r'D:\G\1\Берег'):
#   if i.split('.')[-1] == 'lnk':
#       print i
# # shutil.copy2(src, dst)
# # shutil.copyfile(r'/home/py/mouse.txt', r'/home/py/new-mouse.txt')

# t.Commit()
