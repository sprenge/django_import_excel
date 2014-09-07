#!/usr/bin/env python
import xlrd
import datetime
import random
from bitfield import BitField

from django.contrib.auth.decorators import user_passes_test

from django.db.models.loading import get_model
from  django.contrib.auth.hashers import make_password
from django.core.urlresolvers import reverse_lazy

from django.contrib.auth.models import User, Group

Class ImportExcel :
    def __init__(self):

    wb = xlrd.open_workbook('/tmp/import_excel.xls')
    worksheets = wb.sheet_names()
    ws = wb.sheet_by_name(worksheets[0])

    #process header
    for i in range(0,ws.ncols) :
        cell_value = ws.cell_value(0,i)
        ExcelField2position[cell_value] = i
        Position2Field[i] = cell_value

    print ExcelField2position
    input_ok = True
    for i in range(1, ws.nrows):
        row = []
        for j in range(0,ws.ncols) :
            cell_value = ws.cell_value(i,j)
            row.append(cell_value)
            excel_field = Position2Field[j]
            if ExcelFieldMandatory[excel_field] == True :
                if len(cell_value) < 1 :
                    input_ok = False

    for i in range(1, ws.nrows):
        row = []
        for j in range(0,ws.ncols) :
            cell_value = ws.cell_value(i,j)
            row.append(cell_value)
        for rec in RecordsConstructList :
            appname = rec[0]
            tablename = rec[1]
            m = get_model(appname, tablename)

            kw = {}
            new_table_entry = None
            for name_indexed in rec[2] :
                idx = ExcelField2position[name_indexed[0]]
                val = name_indexed[1](row[idx], row)
                kw[name_indexed[2]] = val
            found = True
            line = None
            try :
                line = m.objects.get(**kw)
            except :
                found = False

            if found == False or rec[5] == True:    #not found and a need to create a record
                new_table_entry = line
                if rec[5] == False :
                    new_table_entry = rec[4](m, row) #create a new one
                #write index first
                for name_indexed in rec[2] :
                    idx = ExcelField2position[name_indexed[0]]
                    val = name_indexed[1](row[idx], row)
                    #print "setattr:", type(new_table_entry),unicode(name_indexed[2]),unicode(val), type(val)
                    setattr (new_table_entry,name_indexed[2],val)
                for name_not_indexed in rec[3] :
                    idx = ExcelField2position[name_not_indexed[0]]
                    val = name_not_indexed[1](row[idx], row)
                    setattr (new_table_entry,name_not_indexed[2],val)
                new_table_entry.save()

    print "input_ok:", input_ok
