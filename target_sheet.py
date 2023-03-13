import win32com.client
from datetime import datetime
import subprocess
from time import sleep
from pyautogui import typewrite,press,click,getWindowsWithTitle


SapGuiAuto = win32com.client.GetObject('SAPGUI')
application = SapGuiAuto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)



dic= [

{'depot':'1310','div':'02','hi':'1','ext':'G'},
{'depot':'1310','div':'02','hi':'2','ext':'G'},
{'depot':'1310','div':'02','hi':'4','ext':'G'},
{'depot':'1310','div':'02','hi':'7','ext':'G'},
{'depot':'1311','div':'02','hi':'1','ext':'G'},
{'depot':'1311','div':'02','hi':'2','ext':'G'},
{'depot':'1311','div':'02','hi':'4','ext':'G'},
{'depot':'1312','div':'02','hi':'1','ext':'G'},
{'depot':'1312','div':'02','hi':'2','ext':'G'},
{'depot':'1312','div':'02','hi':'4','ext':'G'},
{'depot':'1312','div':'02','hi':'7','ext':'G'},
{'depot':'1313','div':'02','hi':'1','ext':'G'},
{'depot':'1313','div':'02','hi':'2','ext':'G'},
{'depot':'1313','div':'02','hi':'4','ext':'G'},
{'depot':'1313','div':'02','hi':'7','ext':'G'},
{'depot':'1314','div':'02','hi':'1','ext':'G'},
{'depot':'1314','div':'02','hi':'2','ext':'G'},
{'depot':'1314','div':'02','hi':'4','ext':'G'},
{'depot':'1314','div':'02','hi':'7','ext':'G'},
{'depot':'1315','div':'02','hi':'1','ext':'G'},
{'depot':'1315','div':'02','hi':'2','ext':'G'},
{'depot':'1315','div':'02','hi':'4','ext':'G'},
{'depot':'1316','div':'02','hi':'1','ext':'G'},
{'depot':'1316','div':'02','hi':'2','ext':'G'},
{'depot':'1316','div':'02','hi':'4','ext':'G'},
{'depot':'1316','div':'02','hi':'7','ext':'G'},
{'depot':'1317','div':'02','hi':'1','ext':'G'},
{'depot':'1317','div':'02','hi':'2','ext':'G'},
{'depot':'1317','div':'02','hi':'4','ext':'G'},
{'depot':'1318','div':'02','hi':'1','ext':'G'},
{'depot':'1318','div':'02','hi':'2','ext':'G'},
{'depot':'1318','div':'02','hi':'4','ext':'G'},
{'depot':'1319','div':'02','hi':'1','ext':'G'},
{'depot':'1319','div':'02','hi':'2','ext':'G'},
{'depot':'1319','div':'02','hi':'4','ext':'G'},
{'depot':'1320','div':'02','hi':'1','ext':'G'},
{'depot':'1320','div':'02','hi':'2','ext':'G'},
{'depot':'1320','div':'02','hi':'4','ext':'G'},
{'depot':'1320','div':'02','hi':'7','ext':'G'},
{'depot':'1321','div':'02','hi':'1','ext':'G'},
{'depot':'1321','div':'02','hi':'2','ext':'G'},
{'depot':'1321','div':'02','hi':'4','ext':'G'},
{'depot':'1321','div':'02','hi':'7','ext':'G'},
{'depot':'1322','div':'02','hi':'1','ext':'G'},
{'depot':'1322','div':'02','hi':'2','ext':'G'},
{'depot':'1322','div':'02','hi':'4','ext':'G'},
{'depot':'1322','div':'02','hi':'7','ext':'G'},
{'depot':'1323','div':'02','hi':'1','ext':'G'},
{'depot':'1323','div':'02','hi':'2','ext':'G'},
{'depot':'1323','div':'02','hi':'4','ext':'G'},
{'depot':'1323','div':'02','hi':'7','ext':'G'},
{'depot':'1324','div':'02','hi':'1','ext':'G'},
{'depot':'1324','div':'02','hi':'2','ext':'G'},
{'depot':'1324','div':'02','hi':'4','ext':'G'},
{'depot':'1325','div':'02','hi':'1','ext':'G'},
{'depot':'1325','div':'02','hi':'2','ext':'G'},
{'depot':'1325','div':'02','hi':'4','ext':'G'},
{'depot':'1326','div':'02','hi':'1','ext':'G'},
{'depot':'1326','div':'02','hi':'2','ext':'G'},
{'depot':'1326','div':'02','hi':'4','ext':'G'},
{'depot':'1327','div':'02','hi':'1','ext':'G'},
{'depot':'1327','div':'02','hi':'2','ext':'G'},
{'depot':'1327','div':'02','hi':'4','ext':'G'},
{'depot':'1328','div':'02','hi':'1','ext':'G'},
{'depot':'1328','div':'02','hi':'2','ext':'G'},
{'depot':'1328','div':'02','hi':'4','ext':'G'},
{'depot':'1329','div':'02','hi':'1','ext':'G'},
{'depot':'1329','div':'02','hi':'2','ext':'G'},
{'depot':'1329','div':'02','hi':'4','ext':'G'},
{'depot':'1329','div':'02','hi':'7','ext':'G'},
{'depot':'1330','div':'02','hi':'1','ext':'G'},
{'depot':'1330','div':'02','hi':'2','ext':'G'},
{'depot':'1330','div':'02','hi':'4','ext':'G'},
{'depot':'1330','div':'02','hi':'7','ext':'G'},
{'depot':'1331','div':'02','hi':'1','ext':'G'},
{'depot':'1331','div':'02','hi':'2','ext':'G'},
{'depot':'1331','div':'02','hi':'4','ext':'G'},
{'depot':'1332','div':'02','hi':'1','ext':'G'},
{'depot':'1332','div':'02','hi':'2','ext':'G'},
{'depot':'1332','div':'02','hi':'4','ext':'G'},
{'depot':'1332','div':'02','hi':'7','ext':'G'},
{'depot':'1333','div':'02','hi':'1','ext':'G'},
{'depot':'1333','div':'02','hi':'2','ext':'G'},
{'depot':'1333','div':'02','hi':'4','ext':'G'},
{'depot':'1333','div':'02','hi':'7','ext':'G'},
{'depot':'1334','div':'02','hi':'1','ext':'G'},
{'depot':'1334','div':'02','hi':'2','ext':'G'},
{'depot':'1334','div':'02','hi':'4','ext':'G'},
{'depot':'1334','div':'02','hi':'7','ext':'G'},
{'depot':'1335','div':'02','hi':'1','ext':'G'},
{'depot':'1335','div':'02','hi':'2','ext':'G'},
{'depot':'1335','div':'02','hi':'4','ext':'G'},
{'depot':'1336','div':'02','hi':'1','ext':'G'},
{'depot':'1336','div':'02','hi':'2','ext':'G'},
{'depot':'1336','div':'02','hi':'4','ext':'G'},
{'depot':'1336','div':'02','hi':'7','ext':'G'},
{'depot':'1310','div':'04','hi':'1','ext':'V'},
{'depot':'1310','div':'04','hi':'2','ext':'V'},
{'depot':'1310','div':'04','hi':'4','ext':'V'},
{'depot':'1310','div':'04','hi':'7','ext':'V'},
{'depot':'1311','div':'04','hi':'1','ext':'V'},
{'depot':'1311','div':'04','hi':'2','ext':'V'},
{'depot':'1311','div':'04','hi':'4','ext':'V'},
{'depot':'1311','div':'04','hi':'7','ext':'V'},
{'depot':'1312','div':'04','hi':'1','ext':'V'},
{'depot':'1312','div':'04','hi':'2','ext':'V'},
{'depot':'1312','div':'04','hi':'4','ext':'V'},
{'depot':'1312','div':'04','hi':'7','ext':'V'},
{'depot':'1313','div':'04','hi':'1','ext':'V'},
{'depot':'1313','div':'04','hi':'2','ext':'V'},
{'depot':'1313','div':'04','hi':'4','ext':'V'},
{'depot':'1313','div':'04','hi':'7','ext':'V'},
{'depot':'1314','div':'04','hi':'1','ext':'V'},
{'depot':'1314','div':'04','hi':'2','ext':'V'},
{'depot':'1314','div':'04','hi':'4','ext':'V'},
{'depot':'1314','div':'04','hi':'7','ext':'V'},
{'depot':'1315','div':'04','hi':'1','ext':'V'},
{'depot':'1315','div':'04','hi':'2','ext':'V'},
{'depot':'1315','div':'04','hi':'4','ext':'V'},
{'depot':'1315','div':'04','hi':'6','ext':'V'},
{'depot':'1316','div':'04','hi':'1','ext':'V'},
{'depot':'1316','div':'04','hi':'2','ext':'V'},
{'depot':'1316','div':'04','hi':'4','ext':'V'},
{'depot':'1316','div':'04','hi':'7','ext':'V'},
{'depot':'1317','div':'04','hi':'1','ext':'V'},
{'depot':'1317','div':'04','hi':'2','ext':'V'},
{'depot':'1317','div':'04','hi':'4','ext':'V'},
{'depot':'1317','div':'04','hi':'6','ext':'V'},
{'depot':'1318','div':'04','hi':'1','ext':'V'},
{'depot':'1318','div':'04','hi':'2','ext':'V'},
{'depot':'1318','div':'04','hi':'4','ext':'V'},
{'depot':'1318','div':'04','hi':'6','ext':'V'},
{'depot':'1319','div':'04','hi':'1','ext':'V'},
{'depot':'1319','div':'04','hi':'2','ext':'V'},
{'depot':'1319','div':'04','hi':'4','ext':'V'},
{'depot':'1319','div':'04','hi':'6','ext':'V'},
{'depot':'1320','div':'04','hi':'1','ext':'V'},
{'depot':'1320','div':'04','hi':'2','ext':'V'},
{'depot':'1320','div':'04','hi':'4','ext':'V'},
{'depot':'1320','div':'04','hi':'6','ext':'V'},
{'depot':'1321','div':'04','hi':'1','ext':'V'},
{'depot':'1321','div':'04','hi':'2','ext':'V'},
{'depot':'1321','div':'04','hi':'4','ext':'V'},
{'depot':'1321','div':'04','hi':'6','ext':'V'},
{'depot':'1322','div':'04','hi':'1','ext':'V'},
{'depot':'1322','div':'04','hi':'2','ext':'V'},
{'depot':'1322','div':'04','hi':'4','ext':'V'},
{'depot':'1322','div':'04','hi':'6','ext':'V'},
{'depot':'1323','div':'04','hi':'1','ext':'V'},
{'depot':'1323','div':'04','hi':'2','ext':'V'},
{'depot':'1323','div':'04','hi':'4','ext':'V'},
{'depot':'1323','div':'04','hi':'6','ext':'V'},
{'depot':'1324','div':'04','hi':'1','ext':'V'},
{'depot':'1324','div':'04','hi':'2','ext':'V'},
{'depot':'1324','div':'04','hi':'4','ext':'V'},
{'depot':'1324','div':'04','hi':'7','ext':'V'},
{'depot':'1325','div':'04','hi':'1','ext':'V'},
{'depot':'1325','div':'04','hi':'2','ext':'V'},
{'depot':'1325','div':'04','hi':'4','ext':'V'},
{'depot':'1325','div':'04','hi':'7','ext':'V'},
{'depot':'1326','div':'04','hi':'1','ext':'V'},
{'depot':'1326','div':'04','hi':'2','ext':'V'},
{'depot':'1326','div':'04','hi':'4','ext':'V'},
{'depot':'1326','div':'04','hi':'7','ext':'V'},
{'depot':'1327','div':'04','hi':'1','ext':'V'},
{'depot':'1327','div':'04','hi':'2','ext':'V'},
{'depot':'1327','div':'04','hi':'4','ext':'V'},
{'depot':'1327','div':'04','hi':'7','ext':'V'},
{'depot':'1328','div':'04','hi':'1','ext':'V'},
{'depot':'1328','div':'04','hi':'2','ext':'V'},
{'depot':'1328','div':'04','hi':'4','ext':'V'},
{'depot':'1328','div':'04','hi':'6','ext':'V'},
{'depot':'1329','div':'04','hi':'1','ext':'V'},
{'depot':'1329','div':'04','hi':'2','ext':'V'},
{'depot':'1329','div':'04','hi':'4','ext':'V'},
{'depot':'1329','div':'04','hi':'7','ext':'V'},
{'depot':'1330','div':'04','hi':'1','ext':'V'},
{'depot':'1330','div':'04','hi':'2','ext':'V'},
{'depot':'1330','div':'04','hi':'4','ext':'V'},
{'depot':'1330','div':'04','hi':'6','ext':'V'},
{'depot':'1331','div':'04','hi':'1','ext':'V'},
{'depot':'1331','div':'04','hi':'2','ext':'V'},
{'depot':'1331','div':'04','hi':'4','ext':'V'},
{'depot':'1331','div':'04','hi':'6','ext':'V'},
{'depot':'1332','div':'04','hi':'1','ext':'V'},
{'depot':'1332','div':'04','hi':'2','ext':'V'},
{'depot':'1332','div':'04','hi':'4','ext':'V'},
{'depot':'1332','div':'04','hi':'7','ext':'V'},
{'depot':'1333','div':'04','hi':'1','ext':'V'},
{'depot':'1333','div':'04','hi':'2','ext':'V'},
{'depot':'1333','div':'04','hi':'4','ext':'V'},
{'depot':'1333','div':'04','hi':'7','ext':'V'},
{'depot':'1334','div':'04','hi':'1','ext':'V'},
{'depot':'1334','div':'04','hi':'2','ext':'V'},
{'depot':'1334','div':'04','hi':'4','ext':'V'},
{'depot':'1334','div':'04','hi':'7','ext':'V'},
{'depot':'1335','div':'04','hi':'1','ext':'V'},
{'depot':'1335','div':'04','hi':'2','ext':'V'},
{'depot':'1335','div':'04','hi':'4','ext':'V'},
{'depot':'1335','div':'04','hi':'6','ext':'V'},
{'depot':'1336','div':'04','hi':'1','ext':'V'},
{'depot':'1336','div':'04','hi':'2','ext':'V'},
{'depot':'1336','div':'04','hi':'4','ext':'V'},
{'depot':'1336','div':'04','hi':'7','ext':'V'}]


s_date ='01.03.2023'
e_date ='31.03.2023'

def division_name(d):
	temp = {'01':'PHARMA','02':'AG','04':'OPHTHA'}
	return temp[d]



def hi_name(h,e):
	temp = {'1':"MIO",'2':"AM",'4':"RSM",'6':"DSM",'7':"SM"}
	if h=='1' and e=='A':
		return temp[h]+'_'+e
	if h=='1' and e=='B':
		return temp[h]+'_'+e
	if h=='1' and e=='AB':
		return temp[h]+'_'+'C'
	return temp[h]
	


# print(hi_name(dic[1]['hi']))



for x in dic:
	i = True
	# session.findById("wnd[0]").maximize()
	session.findById("wnd[0]/usr/ctxtP_DEPOT").text = x['depot']
	session.findById("wnd[0]/usr/ctxtP_DIV").text = x['div']
	session.findById("wnd[0]/usr/ctxtP_EMP_HR").text = x['hi']
	session.findById("wnd[0]/usr/txtP_PGRP").text = x['ext']
	session.findById("wnd[0]/usr/ctxtP_BUDAT-LOW").text = s_date
	session.findById("wnd[0]/usr/ctxtP_BUDAT-HIGH").text = e_date
	session.findById("wnd[0]/usr/ctxtP_BUDAT-HIGH").setFocus()
	session.findById("wnd[0]/usr/ctxtP_BUDAT-HIGH").caretPosition = 10
	session.findById("wnd[0]/tbar[1]/btn[8]").press()
	sleep(3)
	while i:
		a=getWindowsWithTitle("Print")[0]
		if a.title=="Print":
			click(188,75)
			press('enter')
			
			break
	file_name = x['depot']+'_'+hi_name(x['hi'],x['ext'])+'_'+division_name(x['div'])
	# print(file_name)
	sleep(2)
	
	typewrite(file_name)
	press('enter')
	print(file_name)
	# break





