django_import_excel
===================

The goal of this module is to import excel data into multiple tables.  The first row of an excel sheet that will be
imported serves as field identifier.  It is possible to generate multiple database records in different tables from one single
excel data row.

The module expects the following data structures and callback function in order to be able to correctly map excel fields
to database fields.

1. Indication which excel fields are mandatory (set_excelfield_mandatory)

It is a simple dictionary with as key the excel field name and as data True (info is mandatory) or False (info is not mandatory)

2. Field validation rules

Again a dictionary with as key the excel field name and as data a function with the following signature :

callback (self,row,val)

self is the conversion object itself (see later), row is a list with values of the current excel row and val is the value of the
current excel cell.

3. Mapping rules
