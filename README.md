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

The mapping rules have to expressed in the form a list variable with 5 elements :

- First element : path name to application e.g. auth
- Second element : name of the table e.g. User
- Third element : is again a list of two elements, first element is again a list of key mapping elements and ssecond element is a list of non key mapping elements (see further)
- Fourth element : creation callback
- Fifth element : set to False if a record maybe created (if record with the key(s) is not found or to True if a record may not be created. 

mapping elements are always a list of 3 items, the first item is the excel filename, the second element is the callback function to translate excel field to database field and the third element is the field name in the database.

key mapping elements are use to determine whether a record already exists in the database or not (AND function applies if multiple are given)

The callback function takes 3 arguments : self, row (list of values) and val (the value from excel) and return the value as it has to be written into the database.

The create callback takes 3 arguments : self, the model reference and the row (list of values) and has to return the created record instance.  The callback can return None in case of an error.

