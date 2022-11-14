import pandas as pd

filename = 'C:/Users/CHU-TUAN-KIET/Desktop/pythonExcel/dataset.xlsx'


xls = pd.ExcelFile(filename)
lsSheetname = xls.sheet_names
print(f"Tìm thấy {len(lsSheetname)} bảng trong file {filename}")
for i in range(len(lsSheetname)):
    print(f"{i} : {lsSheetname[i]}")

print("Nhập thứ tự insert các bảng \n(Chú ý đến các bảng không có khóa ngoại insert trước, bảng có khóa ngoại insert sau)")

thutu = []

i= 0 
while(i < len(lsSheetname)):
    k = int(input()) 
    if((k) >= len(lsSheetname) or  k < 0):
        print(f"Error : [index] STT phải nhỏ hơn {len(lsSheetname)}")
    else:
        i+=1
        thutu.append(k)


import pypyodbc

conn = pypyodbc.connect('Driver={ODBC Driver 17 for SQL Server};Server=DESKTOP-9RI3QBC;Database=QLTienAn;Trusted_Connection=yes;')

print("Kết nối thành công!!")    

cursor = conn.cursor()


for i in range(len(lsSheetname)):
    sheetname = lsSheetname[thutu[i]]
    # tesst 1 sheet
    df1 = pd.read_excel(xls, sheetname)
    if(not df1.empty):
        
        dicTable = df1.to_dict()

        fields = [key for key in dicTable][1:]

        # LAY DU LIEU TREN HANG 
        strRow = ""
        for i in range(len(dicTable[fields[0]])):
            strline = "("
            for key in fields:
                itemData = dicTable[key][i]
                if(type(itemData) == int):
                    strline += str(itemData) + ", "
                else:
                    strline += "N'" +str(itemData) + "'" + ", "
                    
            strRow += strline[:-2] +"), "
            
        dataInsert = strRow[:-2]

        # LAY TEN COT

        strNameCol = ""
        for i in fields:
            strNameCol += i +','
            
        strNameCol = strNameCol[:-1]

        query = (f"INSERT INTO {sheetname} ({strNameCol}) VALUES {dataInsert}")

        print(query)
        # Bật lên khi cần insert database

        # cursor.execute(query)
        # conn.commit()

        print("Insert Successfully ---------------------------------------------")
    else:
        print("Empty DataFrame")

