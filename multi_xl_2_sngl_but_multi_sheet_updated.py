from pandas import ExcelWriter
import glob
import os
import pandas as pd
from sys import exit

p=input("Please past your excel file's folder path here :  \n\n   : ")
pp=os.chdir(p)
t=len(glob.glob("*.xlsx"))

c=['PLT', 'DHN', 'UTR', 'NRG', 'MYM', 'CHT', 'SYL', 'CUM', 'NKL', 'FEN', 'BGR', 'RAJ', 'RNG', 'DNJ', 'JSR', 'BRS', 'CXB', 'TNG', 'PAB', 'CHP', 'BRB', 'MLB', 'KIS', 'KUS', 'KNJ']

def mrgall():
	
	writer = ExcelWriter("output.xlsx")
	all_data = pd.DataFrame()
	all_data.drop(all_data.index, inplace=True)
	all_data2 = pd.DataFrame()

	for x in range(len(c)):
		for filename in glob.glob("*.xlsx"):
			excel_file = pd.ExcelFile(filename)

			(_, f_name) = os.path.split(filename)
			(f_short_name, _) = os.path.splitext(f_name)

			for sheet_name in excel_file.sheet_names:
				if sheet_name==c[x]:
					df_excel = pd.read_excel(excel_file, sheet_name=sheet_name)
					all_data = all_data.append(df_excel,ignore_index=True)
		all_data.to_excel(writer, c[x], index=False)
		all_data.drop(all_data.index, inplace=True)

	writer.save()
	all_data.drop(all_data.index, inplace=True)

	
mrgall()
