import win32com.client
import subprocess
from time import sleep
from pyautogui import typewrite,press

path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
subprocess.Popen(path)
sleep(3)
sapguiauto = win32com.client.GetObject("SAPGUI")
if not type(sapguiauto)== win32com.client.CDispatch:
	pass
application = sapguiauto.GetScriptingEngine

connection = application.OpenConnection("SAP PRD",True)
sleep(2)

typewrite('UserName')
press('tab')
typewrite('PassWord')
press('enter')