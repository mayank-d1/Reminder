import xlrd
import shutil
import os
import urllib.request
from datetime import datetime,timedelta,date
from tkinter import *
from tkinter import filedialog

canvas_width = 500
canvas_height = 280

master = Tk()
master.title("Silvermine Reminder")
master.geometry("+850+0")

if not os.path.exists("C:/Silvermine_Reminder/Emp_info.xlsx"):  
    os.mkdir("C:/Silvermine_Reminder")
    path = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("excel files","*.xlsx"),("all files","*.*")))
    shutil.copy(path,"C:/Silvermine_Reminder/Emp_info.xlsx")
    urllib.request.urlretrieve("http://2.bp.blogspot.com/-Y7J2iIX0SQs/U-o_9Oe7lUI/AAAAAAABzDc/qTUxiIeF4t4/s1600/Happy%2BBirthday%2BGif%2B1.gif","C:/Silvermine_Reminder/birthday.gif")
    urllib.request.urlretrieve("https://mir-s3-cdn-cf.behance.net/project_modules/disp/b41e1e27075137.5635f8edb514a.gif","C:/Silvermine_Reminder/work.gif")
    img1 = PhotoImage(file="C:/Silvermine_Reminder/birthday.gif")
    img2 = PhotoImage(file="C:/Silvermine_Reminder/work.gif")
    urllib.request.urlretrieve("http://www.iconarchive.com/download/i48622/custom-icon-design/pretty-office-7/Calendar.ico","C:/Silvermine_Reminder/logo.ico")
else:
    path = "C:/Silvermine_Reminder/Emp_info.xlsx"
    img1 = PhotoImage(file="C:/Silvermine_Reminder/birthday.gif")
    img2 = PhotoImage(file="C:/Silvermine_Reminder/work.gif")
    master.wm_iconbitmap("C:/Silvermine_Reminder/logo.ico")

canvas = Canvas(master,width=canvas_width,height=canvas_height)
canvas.pack()

text2 = Text(master, height=10, width=60)

today_date = date.today()

workbook = xlrd.open_workbook(path,on_demand=True)
worksheet = workbook.sheet_by_name("Sheet1")
num_rows = worksheet.nrows
num_cols = worksheet.ncols

person_name = []
previous_birthday={}

def compare_and_insert():
    data = worksheet.cell_value(curr_row, dob_col)
    data = xlrd.xldate_as_tuple(data,0)
    data_date = datetime(*data).date()
    data_date = data_date.replace(year=today_date.year)
    if data_date==today_date:
        name = worksheet.cell_value(curr_row,name_col) # take name from previous column
        person_name.append(name)
    elif data_date >= (today_date-timedelta(days=2)) and data_date<today_date:
        name = worksheet.cell_value(curr_row,name_col) # take name from previous column
        previous_birthday[name]=data_date.strftime("%d-%B-%Y")

for curr_row in range(1, num_rows, 1):
    for curr_col in range(0, num_cols, 1):
        if curr_row == 1:
            data = worksheet.cell_value(curr_row-1, curr_col)

            if data.lower()=="emp_name" or data.lower()=="employee name" or data.lower()=="name":
                name_col = curr_col

            elif data.lower()=="d.o.b." or data.lower()=="dob" or data.lower()=="d.o.b":
                dob_col = curr_col
                compare_and_insert()       
        else:
            compare_and_insert()
            break
    
if person_name:
    canvas.create_image(90,0, anchor=NW, image=img1)
    text2.insert(END,"Today's Birthday:-\n")
    for item in person_name:
        text2.insert(END,item+"\n")
        text2.configure(font="Cambria 11")
        text2.pack()
else:
    canvas.create_image(1,0, anchor=NW, image=img2)
    text2.insert(END,"No Birthday, It's work day!\n\n")
    text2.configure(font="Cambria 11")
    text2.pack()
if len(previous_birthday)!=0:
    text2.insert(END,"\nPrevious Birthday:-\n")
    for item in previous_birthday:
        text2.insert(END,item+"'s birthday was on "+ str(previous_birthday[item])+"\n")
        text2.pack()

mainloop()

