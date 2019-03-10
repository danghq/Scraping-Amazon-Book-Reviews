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

today_date = time.strftime("%Y%B%d")

directory = os.path.expanduser('~\\Desktop\\Amazon Reviews - Downloaded %s' %today_date)

if not os.path.exists(directory):
    os.makedirs(directory)

os.chdir(directory)


opener = urllib2.build_opener()
opener.addheaders = [('User-agent', 'Mozilla\5.0')]

NYTBestSeller = pd.read_excel('C:\Users\Huy\Desktop\NYT Best Sellers.xlsx')

for m in range(0, len(NYTBestSeller['ISBN'])):
    address = 'https://www.amazon.com/product-reviews/'+str(NYTBestSeller['ISBN'][m])+'/ref=cm_cr_arp_d_viewopt_rvwer?ie=UTF8&reviewerType=all_reviews&pageNumber=1'
    url = 'https://www.amazon.com/product-reviews/'+str(NYTBestSeller['ISBN'][m])+'/ref=cm_cr_arp_d_viewopt_rvwer?ie=UTF8&reviewerType=all_reviews&pageNumber='
    response = opener.open(address)
    page = response.read()
    soup = BeautifulSoup(page, 'lxml')
    for button in soup.findAll('li', {'class': 'page-button'}):
        for n in button.findAll('a'):
            max = n.text
            max = int(max.replace(',', ''))


    customer_review = defaultdict(list)

    for n in range(1,max+1):

        url = url+`n`

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

    writer = pd.ExcelWriter("Amazon Book Reviews - " +NYTBestSeller['Title'][m]+".xlsx", engine='xlsxwriter', options={'strings_to_urls': False})
    output.to_excel(writer, sheet_name = 'Amazon Reviews' ,index=False)

    workbook  = writer.book
    worksheet = writer.sheets['Amazon Reviews']
    worksheet.set_column('A:A', 25)
    worksheet.set_column('B:B', 25)
    worksheet.set_column('C:C', 25)
    worksheet.set_column('D:D', 25)
    worksheet.set_column('E:E', 25)
    worksheet.set_column('F:F', 25)

    writer.save()

        