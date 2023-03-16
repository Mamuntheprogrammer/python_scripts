import pandas as pd
import os

print()

direc = input("Please Provide Your File Location : ")

print('''
	** Note **
	1. Your file name must be "cr_list.xlsx"
	2. The Time stamp is Constant so change if you wish to 
	''')

os.chdir(direc)

df = pd.read_excel("cr_list.xlsx", sheet_name=0)
df['Pharma'] = df.apply(lambda x: x['Pharma Due'] + x['Pharma Transit'], axis=1)
df['AG'] = df.apply(lambda x: x['AG Due'] + x['AG Transit'], axis=1)
df['Onco'] = df.apply(lambda x: x['ONCO Due'] + x['ONCO Transit'], axis=1)
df['Ophtha'] = df.apply(lambda x: x['OPTHA Due'] + x['OPTHA Transit'], axis=1)

data = df[['Customer Number','Pharma','AG','Onco','Ophtha']]
data2=pd.pivot_table(data, values=['Pharma', 'AG', 'Onco', 'Ophtha'], index=['Customer Number'], aggfunc=sum)

data3 = data2.reset_index()

Pharma = pd.DataFrame()
AG = pd.DataFrame()
Onco = pd.DataFrame()
Ophtha = pd.DataFrame()


Pharma['Customer']= data3['Customer Number']
Pharma['Credit']= data3['Pharma']
Pharma['Credit segment'] = '1001'


AG['Customer']= data3['Customer Number']
AG['Credit']= data3['AG']
AG['Credit segment'] = '1002'


Onco['Customer']= data3['Customer Number']
Onco['Credit']= data3['Onco']
Onco['Credit segment'] = '1003'

Ophtha['Customer']= data3['Customer Number']
Ophtha['Credit']= data3['Ophtha']
Ophtha['Credit segment'] = '1004'




Master = pd.DataFrame()

Master = pd.concat([Pharma,AG,Onco,Ophtha], axis=0)

Master['Time Stamp']='2023031230302'
Master['Time Stamp1']='2023031230302'
Master['Credit Exp Cat.']='100'
Master['Empty']=''
Master['Amt']='0.00'
Master['Curr'] = 'BDT'

Master = Master[['Customer','Credit segment','Time Stamp','Time Stamp1','Credit Exp Cat.','Empty','Credit','Amt','Curr']]


writer = pd.ExcelWriter("Credit_Exposer2.xlsx", engine='xlsxwriter') 
Master.to_excel(writer, sheet_name='Sheet1',index=False)
writer.save()