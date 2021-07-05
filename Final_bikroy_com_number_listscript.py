import requests
from bs4 import BeautifulSoup
import time 

import mysql.connector
from mysql.connector import Error
from mysql.connector import errorcode


adslink=[]

adsdata=[]

try:
	connection = mysql.connector.connect(host='localhost',
                                         database='bikroy',
                                         user='root',
                                         password='')




	for s in range(1,400):
		src=requests.get("https://bikroy.com/bn/ads?by_paying_member=0&sort=date&order=desc&buy_now=0&page="+str(s)).text

		soup=BeautifulSoup(src,'lxml')

		bdy=soup.find('body')

		#adblock=bdy.find_all('li',class_='normal--2QYVk gtm-normal-ad').a['title']

		for addtitle in bdy.find_all('li',class_='normal--2QYVk gtm-normal-ad'):
			atitle=addtitle.a['href']
			adslink.append(atitle)
		time.sleep(2)

	for ad in adslink:
		src2=requests.get("https://bikroy.com"+ad).text
		soup2=BeautifulSoup(src2,'lxml')

		bdy2=soup2.find('body')
		
		ff=''
		mm=''
		final_data=''
		for address in bdy2.find_all('nav',class_="ui-crumbs"):
			for xx in address.find_all('li','ui-crumb breadcrumb'):
				tt=xx.a.text
				ff=ff+'-'+tt
			#adsdata.append(ff.split('-')[3:7])

		
			for address in bdy2.find_all('li',class_="clearfix"):
				mobile=address.text
				mm=mm+'-'+mobile
			final_data=(ff+' Mobile Numbers '+mm)

			print(final_data)



			mySql_insert_query = """INSERT INTO info (binfo) 
                           VALUES 
                           ('{}') """.format(final_data)

			cursor = connection.cursor()
			result = cursor.execute(mySql_insert_query)
			connection.commit()
			print("Record inserted successfully into Laptop table")
			cursor.close()
		

		time.sleep(1)

except mysql.connector.Error as error:
    print("Failed to insert record into Laptop table {}".format(error))

finally:
    if (connection.is_connected()):
        connection.close()
        print("MySQL connection is closed")


