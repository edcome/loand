# !/usr/bin/python3
# https://stackoverflow.com/questions/14302248/dictionary-update-sequence-element-0-has-length-3-2-is-required
# http://qaru.site/questions/644247/python-help-reading-csv-file-failing-due-to-line-endings
# https://www.tutorialspoint.com/python3/tk_menu.htm
# https://python-scripts.com/import-csv-python
# https://communities.actian.com/s/article/Ingres-ODBC-and-Python
# https://github.com/mkleehammer/pyodbc/wiki/Data-Types
# http://qaru.site/questions/55037/python-how-do-i-know-what-type-of-exception-occurred
# http://www.sqlrelease.com/connecting-python-3-to-sql-server-2017-using-pyodbc

# https://bytes.com/topic/python/answers/649055-pyodbc-data-corruption-problem
# https://github.com/mkleehammer/pyodbc/issues/217
# https://stackoverflow.com/questions/28397527/pyodbc-insert-into-from-a-list
# https://younglinux.info/tkinter/text.php
# 10-aug-2018 add scrollbar to Text widjet

from tkinter import *
from tkinter import filedialog as fd
import os, re, glob
import pyodbc
import csv

# import contextlib

REQ_COLS = 9  # required numbers columns

dbname = r'd:\temp\entire\lob_and\paintdb.accdb'


# dbname = r'c:\ed\Projects\Freelance\Discussed\LobAnd\2\paintDB.accdb'
# dbname = ''
# constr = "DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={0};".format(dbname)

def is_prime(number):
    for element in range(2, number):
        if number % element == 0:
                return False
    return True

def donothing():
    # filewin = Toplevel(root)
    # button = Button(filewin, text="Do nothing button")
    # button.pack()
    for i in range(1, 101):
        text.insert(END, str(i) + "\n")


def prepareFile():
    fname_in = fd.askopenfilename(title="Select file", filetypes=(("CSV files", "*.csv"),))
    if len(fname_in):
        prepare_file(fname_in)


def prepareFolder():
    mypath = fd.askdirectory(title="Select folder")
    if mypath != "":
        mypath += "/"
        only_csv = glob.glob(mypath + "*.csv")
        for i in only_csv:
            prepare_file(i)


def selectBase():
    global dbname
    dbname = fd.askopenfilename(title="Select file", filetypes=(("MS Access files", "*.accdb"),))
    text.insert(END, dbname + "\n")


def importFile():
    global dbname
    if len(dbname):
        fname_in = fd.askopenfilename(title="Select file", filetypes=(("Prepared files", "*.res"),))
        if len(fname_in):
            try:
                constr = "DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={0};".format(dbname)
                dbconn = pyodbc.connect(constr)
                cur = dbconn.cursor()
                import_file(fname_in, cur)
            finally:
                cur.close()
                dbconn.close()

    # if len(dbname):
    # constr = "DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={0};".format(dbname)
    # tbMakeName = create_table(constr,fname_in)
    # if len(tbMakeName):
    # insert_table(constr,fname_in)
    # text.insert(END, 'Connect to ' + dbname + "\n")


def importFolder():
    global dbname
    if len(dbname):
        mypath = fd.askdirectory(title="Select folder")
        if mypath != "":
            mypath += "/"
            only_csv = glob.glob(mypath + "*.res")
            if len(only_csv):
                try:
                    constr = "DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={0};".format(dbname)
                    dbconn = pyodbc.connect(constr)
                    cur = dbconn.cursor()
                    for i in only_csv:
                        import_file(i, cur)
                finally:
                    cur.close()
                    dbconn.close()


def prepare_file(fname_in):
    # messagebox.showinfo("Info", "Process file...")
    text.insert(END, os.path.basename(fname_in) + "\n")
    mypath = os.path.splitext(fname_in)
    fname_out = mypath[0] + '.tmp'
    line = del_shift(fname_in, fname_out)
    if len(line):
        brand_list = get_brand_fields(line)
        brand_list = only_brand_names(brand_list)
        fname_in = fname_out
        mypath = os.path.splitext(fname_in)
        fname_out = mypath[0] + '.res'
        reording_brands(fname_in, fname_out, brand_list)


def del_shift(fname_in, fname_out):
    ptrn = r'&#[^;]*;'
    regexpr = re.compile(ptrn)
    no = 0
    maxcols = 0
    max_line = ""

    with open(fname_in, 'r') as in_file:
        with open(fname_out, 'w') as out_file:
            for line in in_file:
                try:
                    # line = re.sub(regexpr, r'(\1\2\3)', line)
                    line = re.sub(regexpr, my_replace, line)
                    no += 1
                    # count colomns in line
                    cols = line.count(";")
                    if cols > maxcols:
                        maxcols = cols
                        no_max = no
                        # save line
                        max_line = line
                    if no == 1:
                        mincols = cols
                        no_min = no
                    else:
                        if cols < mincols:
                            mincols = cols
                            no_min = no
                except AttributeError:
                    messagebox.showerror("Error", "Line '" + line + "' incorrect.")
                out_file.write(line)
            in_file.close()
            out_file.close()
        if mincols < REQ_COLS:
            messagebox.showerror("Error", "Min columns = '" + str(mincols) + " in line number " + str(no_min))
            return ""
        else:
            return max_line


def my_replace(m):
    if m.group(0) is not None:
        return m.group(0)[:-1]


def get_brand_fields(line):
    lst = line.split(";")
    lst[0:9] = []
    return lst


def only_brand_names(lst):
    for i in range(len(lst)):
        s = lst[i]
        p = s.find(" -")
        lst[i] = s[0:p]
    return lst


def reording_brands(fname_in, fname_out, brand_list):
    no = 0
    with open(fname_in, 'r') as in_file:
        with open(fname_out, 'w') as out_file:
            for line in in_file:
                try:
                    no += 1
                    # replace one brands column on locals columns
                    if no == 1:
                        line = line.replace("brands", ";".join(brand_list))
                        line = line.strip()
                        lst = line.split(";")
                        tmp = lst[7]
                        del lst[7]
                        line = ";".join(lst) + ";" + tmp + "\n"
                    else:
                        line = reorder_line(line, brand_list)
                        if line == "":
                            continue
                except AttributeError:
                    messagebox.showerror("Error", "Line '" + line + "' incorrect.")
                out_file.write(line)
            # if no > 100:
            # break

            in_file.close()
            out_file.close()
        # messagebox.showinfo("Info", "END")
        text.insert(END, "OK.\n")


def reorder_line(line, brand_list):
    line = line.strip()
    lst = line.split(";")
    if len(lst) < REQ_COLS:
        messagebox.showerror("Error", str(len(lst)) + "  " + line)
        return ""

    # if model is empty
    # if lst[3] == "":
    # return ""

    prefix = ";".join(lst[:REQ_COLS])
    lst = lst[REQ_COLS:]

    # if len(lst) == 0:
    # print (line)

    for brand in brand_list:
        prefix += ";"
        if len(lst) > 0:
            for i in range(len(lst)):
                if lst[i].find(brand) >= 0:
                    prefix += lst[i]
                    del lst[i]
                    break

    lst = prefix.split(";")
    tmp = lst[7]
    del lst[7]
    # prefix = ";".join(lst)+tmp + "\n"
    prefix = ";".join(lst) + ";" + tmp + "\n"
    return prefix


def import_file(fname, cur):
    text.insert(END, "import " + os.path.basename(fname) + "\n")
    tbMakeName = create_table(fname, cur)
    if len(tbMakeName):
        insert_table(fname, cur)
        text.insert(END, "import OK.\n")


# text.insert(END, "import " + os.path.basename(fname) + "\n")
# with open(fname, 'r') as csvfile:
# reader = csv.DictReader(csvfile, delimiter=';', quoting=csv.QUOTE_NONE)
# #print (reader.fieldnames)
# num_fields = len(reader.fieldnames)
# text.insert(END, "fields=" + str(len(reader.fieldnames)) + "\n")
# lst = reader.fieldnames[REQ_COLS-1:-1]
# for field in lst:
# text.insert(END, field + "\n")
# for row in reader:
# if reader.line_num == 2:
# # create tb???? table for this make
# tbMakeName = "tb" + row["make"]
# create_table(tbMakeName, lst, cur)
# break
# #text.insert(END, tbMakeName + "\n")
# # else
# # # add record to table
# #text.insert(END, "line No=" + str(reader.line_num) + "     " + row["make"] + "\n")

# csvfile.close()
# insert_table(tbMakeName, fname)
# text.insert(END, "import OK.\n")

# def create_table(tbName, brands, cur):
def create_table(fname, cur):
    tbMakeName = ''
    try:
        with open(fname, 'r') as csvfile:
            reader = csv.DictReader(csvfile, delimiter=';', quoting=csv.QUOTE_NONE)
            # print (reader.fieldnames)
            num_fields = len(reader.fieldnames)
            # text.insert(END, "fields=" + str(len(reader.fieldnames)) + "\n")
            brands = reader.fieldnames[REQ_COLS - 1:-1]
            # for field in brands:
            # text.insert(END, field + "\n")
            for row in reader:
                if reader.line_num == 2:
                    tbMakeName = "tb" + row["make"]
                    tbMakeName = tbMakeName.replace(" ", "")
                    sql = "CREATE TABLE " + tbMakeName + """ (id autoincrement CONSTRAINT MyIDConstraint PRIMARY KEY,
	 [image] varchar(255),
	 year varchar(6),
	 make varchar(50),
	 model varchar(50),
	 paint_color_name varchar(50),
	 code varchar(50),
	 code2 varchar(50),
	 comment varchar(50)"""
                    for brand in brands:
                        brand = brand.replace(" ", "_")
                        sql += ", " + brand + " varchar(50)"
                    sql += ", c memo);"
                    cur.execute(sql)
                    # dbconn.commit()
                    cur.commit()
                    break

    except Exception as ex:
        template = "An exception of type {0} occurred. Arguments:\n{1!r}"
        message = template.format(type(ex).__name__, ex.args)
        print(message)
    finally:
        csvfile.close()
        return tbMakeName


# def insert_table(tbName, data, cur):
# sql = "INSERT INTO " + tbName + """ (id autoincrement CONSTRAINT MyIDConstraint PRIMARY KEY,
# [image] varchar(255),
# year varchar(6),
# make varchar(50),
# model varchar(50),
# paint_color_name varchar(50),
# code varchar(50),
# code2 varchar(50),
# comment varchar(50)"""
# for brand in brands:
# brand = brand.replace(" ","_")
# sql += ", " + brand + " varchar(50)"
# sql += ", c memo);"
# #sql += ");"
# cur.execute(sql)
# #dbconn.commit()
# text.insert(END, sql + "\n")


def insert_table(fname, cur):
    with open(fname, 'r') as csvfile:
        reader = csv.DictReader(csvfile, delimiter=';', quoting=csv.QUOTE_NONE)
        # field_list = []
        input_list = []
        # T = ()
        for row in reader:
            field_list = []
            for field in reader.fieldnames:
                field_list.append(row[field])

            # T = tuple(field_list)
            # input_list.append(T)
            input_list.append(field_list)
            if reader.line_num == 2:
                tbMakeName = "tb" + row["make"]
                tbMakeName = tbMakeName.replace(" ", "")

        for field in reader.fieldnames:
            if field == 'image':
                fields = "[image]"
            else:
                fields += "," + field.replace(" ", "_")
    csvfile.close()

    # sql = "INSERT INTO " + tbMakeName + " (" + fields + ") VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?);"
    sql = "INSERT INTO " + tbMakeName + " (" + fields + ") VALUES(?" + ",?" * (len(reader.fieldnames) - 1) + ");"
    # print(sql)
    cur.executemany(sql, input_list)
    cur.commit()


# with open(fname, 'r') as csvfile:
# reader = csv.DictReader(csvfile, delimiter=';', quoting=csv.QUOTE_NONE)
# #field_list = []
# input_list = []
# #T = ()
# for row in reader:
# field_list = []
# for field in reader.fieldnames:
# field_list.append(row[field])


# # T = tuple(field_list)
# # input_list.append(T)
# input_list.append(field_list)
# if reader.line_num == 2:
# tbMakeName = "tb" + row["make"]

# for field in reader.fieldnames:
# if field == 'image':
# fields = "[image]"
# else:
# fields += "," + field.replace(" ","_")
# csvfile.close()

# sql = "INSERT INTO " + tbMakeName + " (" + fields + ") VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?);"
# #print(input_list[0])
# try:
# dbconn = pyodbc.connect(constr)
# cur = dbconn.cursor()
# cur.executemany(sql,input_list)
# cur.commit()
# except Exception as ex:
# template = "An exception of type {0} occurred. Arguments:\n{1!r}"
# message = template.format(type(ex).__name__, ex.args)
# print (message)
# finally:
# cur.close()
# dbconn.close()


# def del_tmpfiles(mypath):
# if mypath != "":
# mypath += "/"
# only_csv = glob.glob(mypath + "*.tmp")
# for i in only_csv:
# print(i)
# #os.remove(i)

# with contextlib.closing(pyodbc.connect(constr)) as conn:
# sql = "BULK INSERT " + tbName + " FROM " + "'" + fnameCSV + "'" + "WITH ( FIELDTERMINATOR =';', FIRSTROW=2);"
# with contextlib.closing(conn.cursor()) as cursor:
# cursor.execute(sql)
# conn.commit()
# conn.close()

# reader = csv.reader(f, delimiter=';', quoting=csv.QUOTE_NONE)

# for row in reader:
# print(' ### '.join(row))
# # if len(headers) == 0:
# # headers = row
# # for col in row:
# # longest.append(0)
# # type_list.append('')
# # else:
# # for i in range(len(row)):
# # # NA is the csv null value
# # if type_list[i] == 'varchar' or row[i] == 'NA':
# # pass
# # else:
# # var_type = dataType(row[i], type_list[i])
# # type_list[i] = var_type
# # if len(row[i]) > longest[i]:
# # longest[i] = len(row[i])
# f.close()

root = Tk()
menubar = Menu(root)

filemenu = Menu(menubar, tearoff=0)
filemenu.add_command(label="File", command=prepareFile)
filemenu.add_command(label="Folder", command=prepareFolder)
filemenu.add_separator()
filemenu.add_command(label="Exit", command=root.quit)
menubar.add_cascade(label="Prepare", menu=filemenu)

editmenu = Menu(menubar, tearoff=0)
editmenu.add_command(label="Select base", command=selectBase)
editmenu.add_command(label="Import file", command=importFile)
editmenu.add_command(label="Import folder", command=importFolder)
menubar.add_cascade(label="Base", menu=editmenu)

helpmenu = Menu(menubar, tearoff=0)
helpmenu.add_command(label="Help Index", command=donothing)
helpmenu.add_command(label="About...", command=donothing)
menubar.add_cascade(label="Help", menu=helpmenu)

f2 = Frame()
f2.pack()
text = Text(f2, width=50, height=25)
# text.grid(columnspan=2)
text.pack(side=LEFT)

scrollV = Scrollbar(f2, command=text.yview)
# scroll.grid(side=LEFT, fill=Y)
scrollV.pack(side=LEFT, fill=Y)
text.config(yscrollcommand=scrollV.set)

scrollH = Scrollbar(orient=HORIZONTAL, command=text.xview)
# scroll.grid(side=LEFT, fill=Y)
scrollH.pack(side=BOTTOM, fill=X)
text.config(xscrollcommand=scrollH.set)

root.config(menu=menubar)
root.mainloop()