#!/usr/bin/env python
import sys
import traceback
import xlrd
import datetime
from copy import deepcopy

from django.db.models.loading import get_model
from  django.contrib.auth.hashers import make_password
from django.core.urlresolvers import reverse_lazy

from django.contrib.auth.models import User, Group

#1. common callback functions
def nop():
    return [True, '']

def email_or_username(self, row):
    is_email = True
    val = None
    email_idx = self.excelfield2position['email cmsuser']
    username_idx = None
    try :
        username_idx = self.excelfield2position['username']
    except :
        pass

    if username_idx != None :
        #take username
        is_email = False
        val = None
        try :
            val = row[username_idx]
        except :
            pass
        if val == None :
            is_email = True
            val = row[email_idx]
        else :
            if len(val) <= 1 :
                is_email = True
                val = row[email_idx]
    else :
        val = row[email_idx]

    print "eu:", is_email, val
    return [is_email, val]

def check_against_username_email_uniqueness(self):
    email_dict = {}
    username_dict = {}
    for i in range(1, self.ws.nrows):
        row = []
        un = ''
        for j in range(0,self.ws.ncols) :
            cell_value = self.ws.cell_value(i,j)
            excel_field = self.position2field[j]
            if excel_field == 'username' :
                un = cell_value
            row.append(cell_value)

        j = 0
        for cell_value in row:
            excel_field = self.position2field[j]
            if excel_field == 'email cmsuser' :
                email = None
                try :
                    email = User.objects.get(email = cell_value)
                except :
                    pass
                if email_dict.has_key(cell_value) or email != None :
                    if len (un) < 1 :
                        return [False,'Duplicate email and no username specified '+cell_value]
                    else :
                        if username_dict.has_key(un) :
                            return [False,'Duplicate username in excelsheet '+un]
                        else :
                            username = None
                            try :
                                username = User.objects.get(username=un)
                            except :
                                pass
                            if username != None :
                                return [False,'Duplicate username in database '+un]
                            else :
                                username_dict[un] = True
                else :
                    email_dict[cell_value] = True
            j += 1
    return [True,'']

#1.3 common validators
def always_valid(self, row, val):
    return True


#2. Class defintion

class ImportExcel :
    def __init__(self):
        self.precondition_function = nop
        self.excelfield2position = {}
        self.position2field = {}
        self.excelfield_mandatory = {}
        self.record_constructlist = []
        self.path = '/tmp/import_excel.xls'
        self.excelfield_validators = {}
        self.ws = None

    def set_path(self, path):
        self.path = path

    def set_precondition_function(self, precondition_function):
        self.precondition_function = precondition_function

    def set_excelfield_validators(self, validators):
        self.excelfield_validators = deepcopy(validators)

    def set_excelfield_mandatory(self, excelfieldmandatory):
        self.excelfield_mandatory = deepcopy(excelfieldmandatory)

    def set_records_constructlist(self, recordsconstructlist):
        self.record_constructlist = deepcopy(recordsconstructlist)

    def check_mandatory_fields(self):
        pass

    def process_excel_header(self):
        for i in range(0,self.ws.ncols) :
            cell_value = self.ws.cell_value(0,i)
            self.excelfield2position[cell_value] = i
            self.position2field[i] = cell_value


    def run(self):
        '''
        Imports a excel sheet from a fixed location on disk.
        Important is that the workbook contains only one sheet and that the first row contains the header names.
        The mapping between excel field and database in expressed in record_constructlist
        Mandatory fields are indicated in excelfield_mandatory

        Return list with two elements :
        - True (OK) or False (precondition failed)
        - Reason : failure reason (if failed) or number of element imported (OK)
        '''
        failure_reson = ''

        wb = None
        worksheets = None
        ws = None

        try :
            wb = xlrd.open_workbook(self.path)
        except :
            return [False,'Unable to open workbook']
        try :
            worksheets = wb.sheet_names()
        except :
            return [False,'Unable to get sheet names']
        try :
            ws = wb.sheet_by_name(worksheets[0])
        except :
            return [False,'unable to get sheet by name']

        self.ws = ws
        self.process_excel_header()

        input_ok = True
        failure_reason = 'Following excel field is mandatory '
        for i in range(1, ws.nrows):
            row = []
            for j in range(0,ws.ncols) :
                cell_value = ws.cell_value(i,j)
                row.append(cell_value)
                excel_field = self.position2field[j]
                if self.excelfield_mandatory[excel_field] == True :
                    if len(cell_value) < 1 :
                        failure_reason += excel_field
                        input_ok = False
                if self.excelfield_validators.has_key(excel_field) :
                    if self.excelfield_validators[excel_field](self, row, cell_value) == False :
                        return [False, "Invalid field value"+ cell_value+" for field "+excel_field]

        if input_ok == False :
            return [False, failure_reason]

        res = self.precondition_function(self)
        print "result precondition check:", res
        if res[0] != True :
            return res

        for i in range(1, ws.nrows):
            row = []
            #read first the row
            for j in range(0,ws.ncols) :
                cell_value = ws.cell_value(i,j)
                row.append(cell_value)
            #Apply database mapping rules
            for rec in self.record_constructlist :
                appname = rec[0]
                tablename = rec[1]
                m = get_model(appname, tablename)
                kw = {}
                new_table_entry = None
                for name_indexed in rec[2] :
                    idx = self.excelfield2position[name_indexed[0]]
                    val = name_indexed[1](self, row[idx], row)
                    if isinstance(val,list) :
                        return [False,"Invalid field value :"+name_indexed[0]+" row "+str(i+1)+" "+ appname + " " + tablename]
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
                        new_table_entry = None
                        try :
                            new_table_entry = rec[4](self, m, row) #create a new one
                        except :
                            pass
                        if new_table_entry == None :
                            return [False,"record creation failed:"+appname+ " " + tablename+" row "+str(i)]
                    #write index first
                    for name_indexed in rec[2] :
                        idx = self.excelfield2position[name_indexed[0]]
                        val = name_indexed[1](self, row[idx], row)
                        if isinstance(val, list):
                            return [False,"Invalid field value (setattr idx) "+name_indexed[0]+" row "+str(i)+ " "+appname+ " " + tablename]
                        setattr (new_table_entry,name_indexed[2],val)
                    for name_not_indexed in rec[3] :
                        idx = self.excelfield2position[name_not_indexed[0]]
                        val = name_not_indexed[1](self, row[idx], row)
                        if isinstance(val,list) :
                            return [False,"Invalid field value (setattr non idx) "+name_not_indexed[0]+" row "+str(i)+ " " +appname+ " " + tablename]
                        setattr (new_table_entry,name_not_indexed[2],val)
                    save_ok = True
                    try :
                        new_table_entry.save()
                    except :
                        save_ok = None
                    if save_ok == None :
                        type_, value_, traceback_ = sys.exc_info()
                        ex = traceback.format_exception(type_, value_, traceback_)
                        return [False,"Save failed:"+appname+ " " + tablename+" row "+str(i)+ " "+str(ex)+" "+str(sys.exc_value)]
        return [True,""]
