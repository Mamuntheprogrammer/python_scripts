import requests
from bs4 import BeautifulSoup
import time 
adslink=[]

adsdata=[]

for s in range(2,3):
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
	for address in bdy2.find_all('nav',class_="ui-crumbs"):
		for xx in address.find_all('li','ui-crumb breadcrumb'):
			tt=xx.a.text
			ff=ff+'-'+tt
		adsdata.append(ff)
	time.sleep(1)



for i in adsdata:
	print(i)
#print(adblock)