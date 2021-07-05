import pandas as pd
import os 
import datetime
now = datetime.datetime.now()


def fone():
	os.chdir(input("Enter directory :"))
	b=input("Enter file name with extension like book.xlsx  : ")
	a=input("Enter the column name : ")

	df=pd.read_excel(b)
	names = df[a].unique().tolist()

	writer = pd.ExcelWriter("Splited-"+str(now.strftime("%H-%M-%S"))+".xlsx", engine='xlsxwriter')

	for name in names:
	    mydf=df[(df[a] ==name)]
	    mydf.to_excel(writer, sheet_name=str(name),index=False)
	writer.save()
	print("Done ! \n\n")
	
def ftwo():
	os.chdir(input("Enter directory :"))
	b=input("Enter file name with extension like book.xlsx  : ")
	a=input("Enter the 1st column name : ")
	c=input("Enter the 2nd column name : ")

	df=pd.read_excel(b)
	df = df.applymap(str)

	writer = pd.ExcelWriter("Splited.xlsx", engine='xlsxwriter')

	# me=pd.unique(df[['Depot', 'Code']].values.ravel('K'))
	#me=(df['Depot'].append(df['Code'])).unique()

	df2 = df.drop_duplicates(subset=[a,c])

	for i in range(0,len(df2)):
			one=str(df2.iloc[i][a])
			two=str(df2.iloc[i][c])

			t=df[(df[a] == one)&(df[c]==two)]
			#t=df[(df['Depot'] ==depo_code)&(df['RSM Territory']==rsm_code)]
			t.to_excel(writer, sheet_name=one+"-"+two,index=False)
	writer.save()
	print("Done ! \n\n")


def fthree():
	os.chdir(input("Enter directory :"))
	b=input("Enter file name with extension like book.xlsx  : ")
	a=input("Enter the 1st column name : ")
	c=input("Enter the 2nd column name : ")
	d=input("Enter the 3rd column name : ")

	df=pd.read_excel(b)
	df = df.applymap(str)

	writer = pd.ExcelWriter("Splited.xlsx", engine='xlsxwriter')

	# me=pd.unique(df[['Depot', 'Code']].values.ravel('K'))
	#me=(df['Depot'].append(df['Code'])).unique()

	df2 = df.drop_duplicates(subset=[a,c,d])

	for i in range(0,len(df2)):
			one=str(df2.iloc[i][a])
			two=str(df2.iloc[i][c])
			three=str(df2.iloc[i][d])

			t=df[(df[a] == one)&(df[c]==two)&(df[d]==three)]
			#t=df[(df['Depot'] ==depo_code)&(df['RSM Territory']==rsm_code)]
			t.to_excel(writer, sheet_name=one+"-"+two+"-"+three,index=False)
	writer.save()
	print("Done ! \n\n")



while True:
	h=int(input("How many columns you want to filter (max : 3) : "))
	if h==1:
		fone()
	elif h==2:
		ftwo()
	elif h==3:
		fthree()
	else:
		print("Invalid ! try Again :")
