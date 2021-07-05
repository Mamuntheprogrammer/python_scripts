
from selenium import webdriver
import time 
from selenium.webdriver.common.keys import Keys 
import xlsxwriter
# workbook = xlsxwriter.Workbook('Example2.xlsx') 
# worksheet = workbook.add_worksheet()

options = webdriver.ChromeOptions()
options.binary_location = "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe" 
driver = webdriver.Chrome(options=options, executable_path="C:/Users/_007/Desktop/TASK/chromedriver_win32/chromedriver.exe") 


driver.get("https://app.powerbi.com/view?r=eyJrIjoiNjhjZDRkOTgtNzkzNC00MzdmLTgzOTMtYjY1ZmQ3NWNmODc1IiwidCI6ImRmZmRlOTRmLTcyZmItNDlhZS1hY2IyLTBiOTYxYWJkNWI0MSIsImMiOjN9") 

time.sleep(30)
parent = driver.find_element_by_xpath('//*[@id="pvExplorationHost"]/div/div/exploration/div/explore-canvas-modern/div/div[2]/div/div[2]/div[2]/visual-container-repeat/visual-container-modern[1]/transform/div/div[3]/div/visual-modern/div/div/div[2]/div[1]/div[4]/div/div') 


children = parent.find_elements_by_xpath('.//*')

row = 0
column = 0

si=driver.find_elements_by_xpath('//*[@id="pvExplorationHost"]/div/div/exploration/div/explore-canvas-modern/div/div[2]/div/div[2]/div[2]/visual-container-repeat/visual-container-modern[1]/transform/div/div[3]/div/visual-modern/div/div/div[2]/div[1]/div[4]/div/div/div[1]')
print(si)

for child in children:
	print(child.text)
	# t=child.size
	# worksheet.write(row, column, str(child.get_attribute('title')))
	# if t.get('width') == si.get('width'):
	# 	row += 1
	# else:
	# 	row=0
	# 	column += 1
	# 	si['width'] = t.get('width')
	
	
		
# workbook.close()

driver.quit()
# values = [[child.get_attribute('title'),child.size] for child in children]

# print(values)