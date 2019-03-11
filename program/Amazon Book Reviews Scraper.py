import os
import copy
from selenium import webdriver      
from selenium.webdriver.common import action_chains, keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import pandas as pd
import numpy as np
from pattern import web
from collections import defaultdict
import platform 
import time
import urllib
from bs4 import BeautifulSoup
import lxml.html, urllib2, urlparse
import requests
import re
import xlrd

import os

today_date = time.strftime("%Y%B%d")

path = os.path.dirname(os.path.realpath(__file__))

opener = urllib2.build_opener()
opener.addheaders = [('User-agent', 'Mozilla\6.0')]

ISBN = pd.read_excel(path+'\ISBN.xlsx')

for m in range(0, len(ISBN['ISBN'])):

    address = 'https://www.amazon.com/product-reviews/0'+str(ISBN['ISBN'][m])+'/ref=cm_cr_dp_d_acr_sr?ie=UTF8&reviewerType=all_reviews'
    URL1 ='https://www.amazon.com/product-reviews/0'+str(ISBN['ISBN'][m])+'/ref=cm_cr_othr_d_paging_btm_' 
    URL2 = '?ie=UTF8&reviewerType=all_reviews&pageNumber='
    response = opener.open(address)
    page = response.read()
    soup = BeautifulSoup(page, 'lxml')
    for button in soup.findAll('li', {'class': 'page-button'}):
        for n in button.findAll('a'):
            max = n.text
            max = int(max.replace(',', ''))
            

	customer_review = defaultdict(list)

	for n in range(1,448):

		url = URL1+`n`+URL2+`n`
		print url

		opener = urllib2.build_opener()
		opener.addheaders = [('User-agent', 'Mozilla\5.0')]
    
		response = opener.open(url)
		page = response.read()

		soup = BeautifulSoup(page, 'lxml')

		for rating in soup.findAll('div', {'class': 'a-section review'}):
			for i in rating.findAll('a', {'class': 'a-link-normal'}):
				for j in i.findAll(class_='a-icon-alt'):
					j = j.text
					if j=="|":
						pass
					else:
						j = j.encode('utf-8')[0:1]
						j = float(j)
						customer_review['Rating'].append(j)
                    
		for review in soup.findAll(class_='a-section review'):
			for i in review.findAll('span', {'class':'a-size-base review-text'}):
				i = i.text
				customer_review['Review'].append(i)

		for date in soup.findAll('div', {'class': 'a-section review'}):
			for i in date.findAll(class_='a-size-base a-color-secondary review-date'):
				date = i.encode('utf-8').strip('<span class="a-size-base a-color dary review-date">')
				date = date.strip("</span>")
				customer_review['Date'].append(date)

		for author in soup.findAll('div', {'class': 'a-section review'}):
			for i in author.findAll('span', {'class': 'a-size-base a-color-secondary review-byline'}):
				author = i.text.replace('By', '')
				customer_review['Author'].append(author)

		for helpful in soup.findAll('div', {'class': 'a-section review'}):
			for i in helpful.findAll('span', {'class': 'cr-vote-buttons cr-vote-component'}):
				i = i.text
				customer_review['Helpful'].append(i) 

		output=pd.DataFrame(customer_review)

		writer = pd.ExcelWriter(path+"/Amazon Book Reviews - " +ISBN['Title'][m]+".xlsx", engine='xlsxwriter', options={'strings_to_urls': False})
		output.to_excel(writer, sheet_name = 'Amazon Reviews' ,index=False)

		workbook  = writer.book
		worksheet = writer.sheets['Amazon Reviews']


		writer.save()