import tkinter as tk
from tkinter import *
from tkinter import filedialog
from tkinter.ttk import *
import pandas as pd
from tkinter import messagebox
import pypyodbc


global connectstring
global thutu
global lsSheetname
global conn
global isconn
global querystr
global xls



def Handle():
    global xls
    global thutu
    global isconn
    lsSheetname = xls.sheet_names
    index = 0
    result = []
    try:
        if(isconn):
            for i in range(len(lsSheetname)):
                sheetname = lsSheetname[thutu[i]]
                index = i
                # tesst 1 sheet
                df1 = pd.read_excel(xls, sheetname)
                if(not df1.empty):

                    dicTable = df1.to_dict()

                    fields = [key for key in dicTable][1:]
                    print(fields)

                    # LAY DU LIEU TREN HANG
                    strRow = ""
                    for i in range(len(dicTable[fields[0]])):
                        strline = "("
                        for key in fields:
                            itemData = dicTable[key][i]
                            if(type(itemData) == str):
                                strline += "N'" +str(itemData) + "'" + ", "
                            elif(type(itemData) == pd._libs.tslibs.timestamps.Timestamp):
                                convertDate = itemData.to_pydatetime().strftime('%Y-%m-%d')
                                # print(convertDate)
                                strline += "'" +str(convertDate) + "'" + ", "
                            elif(type(itemData) == float):
                                # print(type(itemData))
                                import numpy as np
                                #https://note.nkmk.me/en/python-nan-usage/
                                if(np.isnan(itemData)):
                                    convertInt = "null"
                                else:
                                    convertInt = int(itemData)
                                strline += str(convertInt) + ", "

                            else:
                                # print(type(itemData), itemData)
                                strline += str(itemData) + ", "

                        strRow += strline[:-2] +"), "

                    dataInsert = strRow[:-2]

                    # LAY TEN COT

                    strNameCol = ""
                    for i in fields:
                        strNameCol += i +','

                    strNameCol = strNameCol[:-1]

                    query = (f"INSERT INTO {sheetname} ({strNameCol}) VALUES {dataInsert}")
                    # print(query)
                    result.append(query)

                else:
                    print(f"{sheetname} is Empty DataFrame")

            return result
    except Exception as e:
        messagebox.showerror('Error', f"{e} - {index}")

def GetListSheetName():
    global xls

    lsSheetname = xls.sheet_names
    result = ""
    for i in range(len(lsSheetname)):
        result += f"| {i} : {lsSheetname[i]} |\n "
    return result


connectstring = "Driver={ODBC Driver 17 for SQL Server};Server=DESKTOP-9RI3QBC;Database=QLTienAn2;Trusted_Connection=yes;"
def openFile():
    filepath = filedialog.askopenfilename(initialdir="",
                                          title="Open file okay?",
                                          filetypes= (("text files","*"),
                                          ("all files","*.*")))
    txt.set(filepath)


def connSQL():
    global conn
    global isconn
    global connectstring
    connectstring = t1.get("1.0",END)
    try:
        conn = pypyodbc.connect(connectstring)
        messagebox.showinfo("OK","K???t n???i th??nh c??ng!!")
        isconn = True
    except Exception as e:
        isconn = False
        messagebox.showerror("???nn't","Ki???m tra l???i, ?????c h?????ng d???n t???o chu???i nh??!")



def CheckFile():
    global lsSheetname
    global xls

    filepath = e1.get()
    xls = pd.ExcelFile(filepath)
    lsSheetname = xls.sheet_names
    lbmsg = tk.Label(text=f"T??m th???y {len(lsSheetname)} b???ng trong file {filepath}")
    lbmsg.grid(row = 40, column = 25, sticky = W, pady = 2)

    lbmsg = tk.Label(text=GetListSheetName(), bg='#1A9F61')


    lbmsg.grid(row = 50, column = 25, sticky = W, pady = 2)

    lbmsg = tk.Label(text=f"Nh???p th??? t??? insert c??c b???ng (gi?? tr??? ph???i nh??? h??n {len(lsSheetname)}) \n(Ch?? ?? ?????n c??c b???ng kh??ng c?? kh??a ngo???i insert tr?????c, b???ng c?? kh??a ngo???i insert sau) (Ph??n c??ch nhau b???i d???u ph???y. VD: 0,1,3,4,5)", fg='#ff0000')
    lbmsg.grid(row = 60, column = 25, sticky = W, pady = 2)


    #---------------------



def checkThuTu():
    global thutu
    thutu = []
    if("," in e2.get()):
        try:
            thutu = e2.get().split(",")

            thutu = [int(i) for i in thutu]
            print(thutu)
            
            if(len(thutu) < len(lsSheetname)):
                messagebox.showerror("???nn't",f"B???n ph???i nh???p ????? s??? l?????ng b???ng")
                return
            for i in thutu:
                if(i >= len(lsSheetname)):
                    messagebox.showerror("???nn't",f"STT ph???i nh??? h??n {len(lsSheetname)}")
                    return
            
                
            
            messagebox.showinfo("OK","???n b???n ??i")
        except Exception as e:
            messagebox.showerror("Error",e)

    else:
        messagebox.showerror("???nn't","Ph???i nh???p chu???i c?? d???u ph???y!")


def createQuery():
    global xls
    global thutu
    global querystr
    querystr = Handle()
    for i in querystr:
        t2.insert(tk.END,i+";\n")



def InsertToDB():
    global conn
    global isconn
    global querystr
    if(isconn):
        try:
            cursor = conn.cursor()
            query = ""
            for i in querystr:
                # print(i)
                query += (i) +";"
            print(query)
            cursor.execute(query)
            conn.commit()
            messagebox.showinfo("OK","Th??m th??nh c??ng")
        except Exception as e:
            messagebox.showerror("Error",e)




window = tk.Tk()


window.title("CTK TOOL C???C M???NH VIP PRO ")
window.geometry('1000x600')

txt = StringVar()
txt1 = StringVar()

label = tk.Label(text="Tool Insert Dataset from Excel to MSSQL by CTK", font=20,border=5, fg="#6366F1")
label.grid(row = 0, column = 25, sticky = W, pady = 2)

button = Button(text="Ch???n file data (xlxs)",command=openFile, width=60)
button.grid(row = 20, column = 25, sticky = W, columnspan = 2)

label1 = tk.Label(text="???????ng d???n ?????n file Excel : ")
label1.grid(row = 30, column = 0, sticky = W, pady = 2)

e1 = tk.Entry(window,width=60,textvariable=txt)
e1.grid(row = 30, column = 25, sticky = W, pady = 2)

button = Button(text="Submit",command=CheckFile)
button.grid(row = 30, column = 60, sticky = W, columnspan = 2)


#------------------------------------------
e2 = tk.Entry(window,width=60)
e2.grid(row = 80, column = 25, sticky = W, pady = 2)

btn1 = Button(text="Check",command=checkThuTu)
btn1.grid(row = 80, column = 60, sticky = W, columnspan = 2)
#------------------------------------------
label1 = tk.Label(text="S???a connection str: ")
label1.grid(row = 90, column = 0, sticky = W, pady = 2)

t1 = tk.Text(window,width=45,height=5)
t1.insert(tk.END,connectstring)
t1.grid(row = 90, column = 25, sticky = W, pady = 2)

btn3 = Button(text="Connect",command=connSQL)
btn3.grid(row = 90, column = 60, sticky = W, columnspan = 2)

#------------------------------------------
btn2 = Button(text="T???o SQL Query", width=60,command=createQuery)
btn2.grid(row = 100, column = 25, sticky = W, columnspan = 2)

#------------------------------------------


l2 = tk.Label(text="K???t qu???:chu???i SQL query ")
l2.grid(row = 110, column = 0, sticky = W, pady = 2)

t2 = tk.Text(window,width=45,height=10)
t2.grid(row = 110, column = 25, sticky = W, pady = 2)

btn2 = Button(text="Insert Query to DB",command=InsertToDB, width=60)
btn2.grid(row = 120, column = 25, sticky = W, columnspan = 2)



#------------------------------------------





window.mainloop()

