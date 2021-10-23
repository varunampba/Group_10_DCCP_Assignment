# -*- coding: utf-8 -*-
"""
Created on Sat Oct 16 12:40:31 2021

"""
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
import pandas as pd
from time import localtime, strftime
import time

# Opening the Website
path = 'C:/Users/Jaspreet Singh/Desktop/ISB/2. Term-1/Data Collection and Preprocessing/2. Assignment/chromedriver.exe' # CHANGE THIS PATH 
driver = webdriver.Chrome(path) # selecting the browser to be opened
driver.get("https://nasscom.in/members-listing") # selecting the website

time.sleep(1)

Company_Name = []
City_Name = []
Website = []

# Page - 1
company_name = driver.find_elements_by_css_selector("div[class='views-field views-field-title']")
city_name = driver.find_elements_by_css_selector("div[class='views-field views-field-field-city-members-list']")
website = driver.find_elements_by_css_selector("div[class='views-field views-field-field-website-member']")

for company in company_name:
    Company_Name.append(company.text)
    
for city in city_name:
    City_Name.append(city.text)
    
for web in website:
    fc = web.find_element_by_css_selector("div[class='field-content']")
    url = fc.find_element_by_tag_name("a")
    Website.append(url.get_attribute('href'))



# Rest of the pages
for pages in range(1,10):
    
    next_page = driver.find_element_by_tag_name("li[class='pager-next']")
    next_page.click()
    time.sleep(1)
    
    company_name = driver.find_elements_by_css_selector("div[class='views-field views-field-title']")
    city_name = driver.find_elements_by_css_selector("div[class='views-field views-field-field-city-members-list']")
    website = driver.find_elements_by_css_selector("div[class='views-field views-field-field-website-member']")

    for company in company_name:
        Company_Name.append(company.text)
    
    for city in city_name:
        City_Name.append(city.text)
    
    for web in website:
        fc = web.find_element_by_css_selector("div[class='field-content']")
        url = fc.find_element_by_tag_name("a")
        Website.append(url.get_attribute('href'))
    

Output_DF = pd.DataFrame(
    {'Company_Name': Company_Name,
     'City_Name': City_Name,
     'Website': Website
     })

time.sleep(1)

datatoexcel = pd.ExcelWriter('Nasscom_Members_123.xlsx')    
Output_DF.to_excel(datatoexcel)
datatoexcel.save()

time.sleep(1)
driver.close()