import pandas as pd

#Read the excel sheet to pandas dataframe

zmonachv = pd.read_excel("zmonachv.xlsx", sheet_name="Sheet1")

temp = pd.read_excel("hierarchy.xlsx", sheet_name=0)

hierarchy = temp[['Employee', 'Plant','Level.8','Industry code 1']]
hierarchy.columns = ['Customer','Plant','AGM','Industry code 1']

name = pd.read_excel("depot.xlsx", sheet_name="name")
depot = pd.read_excel("depot.xlsx", sheet_name="depot")

temp = pd.merge(hierarchy,name, on='AGM',how='inner')
temp2 = pd.merge(temp,depot,on='Plant',how='inner')
final = pd.merge(zmonachv,temp2,on='Industry code 1',how='left')

final['Month'] = final['Start Date'].dt.month_name().str[:3]
final['Title'] = final['Designation'].fillna('')+' Name: '+final['Name 1']+' Depot: '+final['Depot-NAME']+ ' (Territory:'+final['Industry code 1']+' ID Code: '+final['Customer_x'].str[-5:]+')'

writer = pd.ExcelWriter("Master_Data.xlsx", engine='xlsxwriter') 
final.to_excel(writer, sheet_name='Sheet1',index=False)
writer.save()