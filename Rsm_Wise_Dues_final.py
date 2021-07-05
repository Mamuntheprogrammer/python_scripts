import pandas as pd
import os
import numpy as np

a=input("Enter Folder Path  : ")
b=input("Chemist Wise File Name : ")
c=input("RSM Wise File Name : ")


#os.chdir(r'E:\REPORT_BUILDER\Final')
os.chdir(a)

df=pd.read_excel(b+".XLSX")
#df=pd.read_excel(r"Chemist-Wise-Pharma.XLSX")

#df2=pd.read_excel(r"Rsm-Wise-Pharma.XLSX")
df2=pd.read_excel(c+".XLSX")

for i in range(0,len(df2)):
	depo_code=int(df2.iloc[i]['Depot'])
	rsm_code=str(df2.iloc[i]['Employee Territory'])


	t=df[(df['Depot'] ==depo_code)&(df['RSM Territory']==rsm_code)]

	

	f=t[['Customer Code','Customer Name','Territory','MIO ID','MIO Name','AM/SAM ID','AM/SAM Ter','AM/SAM Name','Dues'
]]
	
	
	#f.at['Total','Dues']=f['Dues'].sum()


	f.sort_values(by='AM/SAM Ter' , inplace=True)

	

	f.loc[len(f)+1]=['','','','','','','','TOTAL',f['Dues'].sum()]

	#f.loc[len(f)+2]=[f['A'],f['RSM Name'],f['RSM Territory'],'','','','','','']
	
	
	f.to_excel(str(depo_code)+"-"+rsm_code+".xlsx",index=False)
	













'''

for i in range(0,len(df2)):
	print(df2['Depot'])


#df2.to_excel("output.xlsx",index=False)


t=df[(df['DEPOT'] ==1335)&(df['TERR']=='KURSM1P')]
#t.to_excel("output.xlsx",index=False)
'''