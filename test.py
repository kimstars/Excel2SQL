import tkinter as tk
from tkinter import *
from tkinter import messagebox
window = tk.Tk()
window.title("CTK TOOL CỰC MẠNH VIP PRO ")
window.geometry('1000x500')

def checkThuTu():
    if("," in e2.get()):
        thutu = e2.get().split(",")
        print(thutu)
        
        thutu = [int(i) for i in thutu]
        for i in thutu:
            if(i >= 4):
                messagebox.showerror("Error","STT phải nhỏ hơn 4")
                break
    else:
        messagebox.showerror("Error","Phải nhập chuỗi có dấu phẩy!")
        

e2 = tk.Entry(window,width=40)
e2.grid(row = 30, column = 25, sticky = W, pady = 2)

button = Button(text="Submit",command=checkThuTu)
button.grid(row = 30, column = 42, sticky = W, columnspan = 2)
window.mainloop()
