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
import os
import re
import glob

import pyodbc
import csv

import configparser

from tkinter import messagebox

import msaccessdb

# import contextlib

CONFIG_FILE = "settings.ini"
REQ_COLS = 9  # required numbers columns

#dbname = r'd:\temp\entire\lob_and\paintdb.accdb'



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
    # global dbname
    # dbname = fd.askopenfilename(title="Select file", filetypes=(("MS Access files", "*.accdb"),))
    base_path = fd.askopenfilename(title="Select file", filetypes=(("MS Access files", "*.accdb"),))
    if len(base_path) > 0:
        config["BASE"]["BASE_PATH"] = base_path
    label1.config(text=config["BASE"]["BASE_PATH"])


def importFile():
    base_path = config["BASE"]["BASE_PATH"]
    if len(base_path):
        fname_in = fd.askopenfilename(title="Select file", filetypes=(("Prepared files", "*.res"),))
        if len(fname_in):
            try:
                constr = "DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={0};".format(base_path)
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
    base_path = config["BASE"]["BASE_PATH"]
    if len(base_path):
        mypath = fd.askdirectory(title="Select folder")
        if mypath != "":
            mypath += "/"
            only_csv = glob.glob(mypath + "*.res")
            if len(only_csv):
                try:
                    constr = "DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={0};".format(base_path)
                    dbconn = pyodbc.connect(constr)
                    cur = dbconn.cursor()
                    for i in only_csv:
                        import_file(i, cur)
                finally:
                    cur.close()
                    dbconn.close()


def prepare_file(fname_in):
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

    # ignore lines where model field is empty
    if lst[3] == "":
        return ""

    prefix = ";".join(lst[:REQ_COLS])
    lst = lst[REQ_COLS:]

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
        createIndex(tbMakeName, cur)
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

def createIndex(tbMakeName, cur):
    sql = "CREATE INDEX " + tbMakeName + "Index ON " + tbMakeName + " (code);"
    cur.execute(sql)
    # dbconn.commit()
    cur.commit()
# CREATE INDEX индекс
# ON таблица (поле [ASC|DESC][, поле [ASC|DESC], ...])
# [WITH { PRIMARY | DISALLOW NULL | IGNORE NULL }]


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



def getUniquePaintCode(cur):
    sql = "SELECT DISTINCT code FROM tbAcura;"
    cur.execute(sql)
    return cur.fetchall()

def getCarsInPaintCode(code, cur):
    #TODO: move select field 'make' to another place. Select only one record !!!
    sql = "SELECT model, paint_color_name, year, make FROM tbAcura WHERE code='" + code + "';"
    cur.execute(sql)
    return cur.fetchall()

def makeHeaders():
    base_path = config["BASE"]["BASE_PATH"]
    if len(base_path):
        try:
            constr = "DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={0};".format(base_path)
            dbconn = pyodbc.connect(constr)
            cur = dbconn.cursor()
            header_1 = {}
            header_2 = {}
            codes = getUniquePaintCode(cur)
            # text.insert(END, type(codes))
            # text.insert(END, "\n")
            for code in codes:
                models = {}
                cars = getCarsInPaintCode(code[0], cur)
                for car in cars:
                    if not car[0] in models:
                        models[car[0]] = []
                    models[car[0]].append(car[2])

                years = get_unique_years(models)
                years.sort()
                header = car[1] + " " + code[0] + " <> " + " ".join(years)

                # add only last two digits of the year
                for year in years:
                    header += " " + year[-2:]

                # add make UPCASE
                header += " " + car[3].upper()

                # add models
                for key in models.keys():
                    header += " " + key
                header_1[code[0]] = header
                # text.insert(END, str(len(header_1)))
                # text.insert(END, "\n")

                header = makeHeader_2(code[0], models)
                header_2[code[0]] = header
                # text.insert(END, header)
                # text.insert(END, "\n")



            # paint_color_name

            # text.insert(END, header)
            # text.insert(END, "\n")

            # csv = make_result_csv(header_1, header_2, cur)
            # text.insert(END, csv)
            #     text.insert(END, "\n")
            make_result_csv(header_1, header_2, cur)

            messagebox.showinfo("Info", str(len(codes)))
        # except:
        #     pass
        finally:
            cur.close()
            dbconn.close()


def get_unique_years(models):
    res = set()
    for years in models.values():
        res |= set(years)

    return list(res)

def makeHeader_2(code, models):
    used_keys = []
    header = ""
    for key in models.keys():
        if not key in used_keys:
            s_years = set(models[key])
            used_keys.append(key)
            res = find_equ(models, used_keys, s_years)
            if len(header) > 0:
                header += " , "
            header += " ".join(models[key]) + " " + key
            if len(res) > 0:
                header += " " + " ".join(res)
                used_keys += res
    return header

def find_equ(models, used, s_years):
    res = []
    for key in models.keys():
        if not key in used:
            if s_years == set(models[key]):
                res.append(key)
    return res


def make_result_csv(h1, h2, cur):
    res = []
    sql = "SELECT * FROM tbAcura;"
    cur.execute(sql)
    records = cur.fetchall()
    try:
        with open("acura_heads.csv", 'w') as csvfile:
            #res.append(';'.join(t[0] for t in cur.description) + ";header_1;header_2")
            csvfile.write(';'.join(t[0] for t in cur.description) + ";header_1;header_2" + "\n")
            #print(res)
            for rec in records:
                rec[0] = str(rec[0])
                # res.append(h1[rec[6]])
                # res.append(h2[rec[6]])
                csv_line = ";".join(rec) + ";" + h1[rec[6]] + ";" + h2[rec[6]] + "\n"
                csvfile.write(csv_line)
                # res.append(csv_line)
            # return res
    finally:
        csvfile.close()

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




def callback():
    #if messagebox.askokcancel("Quit", "Do you really wish to quit?"):
    config.write(open(conf_path, "w"))
    root.destroy()


# read settings
conf_path = os.path.join(os.path.dirname(__file__), CONFIG_FILE)
config = configparser.ConfigParser()
config.read(conf_path)

if not config.has_section("BASE"):
    config.add_section("BASE")
    config.set("BASE","BASE_PATH", "")

# try:
#     base_path = config["BASE"]["BASE_PATH"]
# except KeyError:
#     base_path = "???"



class MyDialog:
    def __init__(self, parent):
        top = self.top = Toplevel(parent)
        self.myLabel = Label(top, text='Enter file name below')
        self.myLabel.pack()

        self.myEntryBox = Entry(top)
        self.myEntryBox.pack()
        self.mySubmitButton = Button(top, text='Ok', command=self.send)
        self.mySubmitButton.pack()

    def send(self):
        global username
        username = self.myEntryBox.get()
        self.top.destroy()


def createBase():
    # Create empty .accdb file
    inputDialog = MyDialog(root)
    root.wait_window(inputDialog.top)
    print('Username: ', username)
    #msaccessdb.create("new_paintdb.accdb")

root = Tk()
root.title("Service paintDB")
# root['bg'] = '#aa00aa'
# root.state('zoomed')
menubar = Menu(root)

filemenu = Menu(menubar, tearoff=0)
filemenu.add_command(label="File", command=prepareFile)
filemenu.add_command(label="Folder", command=prepareFolder)
filemenu.add_separator()
filemenu.add_command(label="Exit", command=root.quit)
menubar.add_cascade(label="Prepare", menu=filemenu)

editmenu = Menu(menubar, tearoff=0)
editmenu.add_command(label="Create base", command=createBase)
editmenu.add_command(label="Select base", command=selectBase)
editmenu.add_command(label="Import file", command=importFile)
editmenu.add_command(label="Import folder", command=importFolder)
menubar.add_cascade(label="Base", menu=editmenu)


helpmenu = Menu(menubar, tearoff=0)
helpmenu.add_command(label="Open options", command=donothing)
helpmenu.add_command(label="Make headers", command=makeHeaders)
menubar.add_cascade(label="Options", menu=helpmenu)


helpmenu = Menu(menubar, tearoff=0)
helpmenu.add_command(label="Help Index", command=donothing)
helpmenu.add_command(label="About...", command=donothing)
menubar.add_cascade(label="Help", menu=helpmenu)

f1 = Frame()
f1.pack(side=TOP, expand=YES, fill=X)
#label1 = Label(f1, width=80, text=config["BASE"]["BASE_PATH"], fg="#eee", bg='#0a0a0a')
# label1 = Label(f1, width=80, textvariable=config["BASE"]["BASE_PATH"])
label1 = Label(f1, width=80, text=config["BASE"]["BASE_PATH"])

label1.pack(expand=YES, fill=X)
# label1.pack(side=BOTTOM)

f2 = Frame()
f2.pack()
text = Text(f2, width=80, height=25)
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




root.protocol("WM_DELETE_WINDOW", callback)


root.mainloop()