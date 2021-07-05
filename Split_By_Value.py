import pandas as pd
import os 

os.chdir(input("Enter directory :"))
b=input("Enter file name with extension like book.xlsx  : ")
a=input("Enter the column name : ")

df=pd.read_excel(b)
names = df[a].unique().tolist()
writer = pd.ExcelWriter("Splited.xlsx", engine='xlsxwriter')
for myname in names:
    mydf = df.loc[df.Depot==myname]
    mydf.to_excel(writer, sheet_name=myname,index=False)

writer.save()