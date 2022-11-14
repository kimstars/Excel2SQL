# Excel2SQL 

Excel2SQL  is a Tool convert Excel dataset to SQL query and insert it to your database 

## How to use

- install requirements : pip install -r requirements.txt
- run : python Excel2SQL.py



## Examples

![image-20221114214251830](C:\Users\CHU-TUAN-KIET\AppData\Roaming\Typora\typora-user-images\image-20221114214251830.png)



- <u>**Step 1 : Select dataset (xlxs)**</u>

[Required!]

You must prepare your dataset Excel 

with **sheetname** = **tablename** and **name_column** = **fieldname**

Click **submit**

- <u>**Step 2 : Input your order to read the tables**</u>

To avoid conflict among **foreign_keys** of tables

- <u>**Step 3 : Edit your connection string**</u>

Driver={**ODBC Driver 17 for SQL Server**};Server=**DESKTOP**;Database=**DBname**;Trusted_Connection=yes;

ODBC Driver 17 for SQL Server : search version name ODBC on your PC

DESKTOP : username of server

DBname : your DB name

And click **connect**

- Step 4 : Auto genarate your SQL query
- Step 5 : Insert SQL query to your DB

