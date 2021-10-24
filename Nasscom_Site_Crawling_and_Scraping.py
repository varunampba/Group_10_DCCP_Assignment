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

time.sleep(1.5)
ignored_exceptions=(NoSuchElementException,StaleElementReferenceException)

Company_Name = []
City_Name = []
Website = []

CIN = []
State = []
Incorpoation_Date = []

Company_Legal_Name = []
Company_Type = []
Authorised_Capital = []
Paid_up_Capital = []
Date_of_AGM = []
Date_of_Balance_Sheet = []
Industry = []
Segment = []
Website_1 = []
Key_Person_Name = []
Current_Loan_Amount = []
Loan_Amount_Satisfied = []
Top_Lender_Name = []
Total_Lenders = []
Last_Charge_Activity = []

# Page - 1
company_name = driver.find_elements_by_css_selector("div[class='views-field views-field-title']")
city_name = driver.find_elements_by_css_selector("div[class='views-field views-field-field-city-members-list']")
website = driver.find_elements_by_css_selector("div[class='views-field views-field-field-website-member']")

for company in company_name:
    Company_Name.append(company.text)
    
for city in city_name:
    City_Name.append(city.text)
    
for web in website:
    fc = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='field-content']"))
        )
    try:
        url = fc.find_element_by_tag_name("a")
        Website.append(url.get_attribute('href'))
    except:
        Website.append('NA')
    
    
# Rest of the pages
for pages in range(1,270):
    
    next_page = driver.find_element_by_tag_name("li[class='pager-next']")
    next_page.click()
    time.sleep(1.5)
    
    company_name = driver.find_elements_by_css_selector("div[class='views-field views-field-title']")
    time.sleep(0.25)
    for company in company_name:
        try:
            Company_Name.append(company.text)
        except:
            Company_Name.append("Stale_Error")
            print("Stale Error in Company_Name")
            
    city_name = driver.find_elements_by_css_selector("div[class='views-field views-field-field-city-members-list']")
    time.sleep(0.25)
    for city in city_name:
        try:
            City_Name.append(city.text)
        except:
            City_Name.append("Stale_Error")
            print("Stale Error in City_Name")
            
    website = driver.find_elements_by_css_selector("div[class='views-field views-field-field-website-member']")
    time.sleep(0.25)
    for web in website:
        try:
            fc = web.find_element_by_css_selector("div[class='field-content']")
        except:
            print("Stale Error in Website")
            pass
        try:
            url = fc.find_element_by_tag_name("a")
            Website.append(url.get_attribute('href'))
        except:
            Website.append('NA')


Total_Companies = len(Company_Name)   
print("Total Companies Scraped from Nasscom Site = "+ str(Total_Companies))

# Getting CIN from MCA
driver.get("https://www.mca.gov.in/mcafoportal/showCheckCompanyName.do")
time.sleep(1)

co_num = 0

for co in Company_Name:
    co_num = co_num + 1
    
    co_list = co.split(' ')
    if len(co_list) == 0:
        word = "Not_Found"
    elif len(co_list) == 1: 
        word = co_list[0]
    elif len(co_list) <= 4:
        word = co_list[0]+' '+co_list[1]
    else:
        word = co_list[0]+' '+co_list[1]+' '+co_list[2]
        
    search = driver.find_element_by_id("name1")
    search.send_keys(Keys.CONTROL, 'a')
    search.send_keys(Keys.BACKSPACE)
    search.send_keys(word)
    search.send_keys(Keys.RETURN) 
    
    try:
        table_rows = driver.find_elements_by_css_selector("td[align='center']")
        CIN.append(table_rows[0].text)
        State.append(table_rows[2].text)
        Incorpoation_Date.append(table_rows[3].text)
        print("MCA Scraping Progress: "+str(co_num)+" / "+str(Total_Companies))
    except:
        time.sleep(1)
        cross = driver.find_element_by_id("msgboxclose")
        cross.click()
        CIN.append("Company not found on MCA")
        State.append(" ")
        Incorpoation_Date.append("Company not found on MCA")
        print("MCA Scraping Progress: "+str(co_num)+" / "+str(Total_Companies)+" Not found in MCA")

print(" ")

# Company financial Details from the third party website connected with MCA API
driver.get("https://www.thecompanycheck.com/company/")
time.sleep(3)

co_num = 0

for num in CIN:
    co_num = co_num + 1
    
    search_1 = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
            EC.presence_of_element_located((By.ID, "txtSearch"))
            )
    search_1.send_keys(Keys.CONTROL, 'a')
    search_1.send_keys(Keys.BACKSPACE)
    search_1.send_keys(num)
    search_button = WebDriverWait(driver, 0, ignored_exceptions=ignored_exceptions).until(
        EC.presence_of_element_located((By.ID, "btnSearch"))
        )
    search_button.click()
    
    if num == "Company not found on MCA" or "-" in num or 'F0' in num:
        Company_Legal_Name.append("NA in MCA")    
        Company_Type.append("NA in MCA")   
        Authorised_Capital.append("NA in MCA")
        Paid_up_Capital.append("NA in MCA")
        Date_of_AGM.append("NA in MCA")
        Date_of_Balance_Sheet.append("NA in MCA")
        Industry.append("NA in MCA")
        Segment.append("NA in MCA")
        Website_1.append("NA in MCA")
        Key_Person_Name.append("NA in MCA")
        Current_Loan_Amount.append("NA in MCA")
        Loan_Amount_Satisfied.append("NA in MCA")
        Top_Lender_Name.append("NA in MCA")
        Total_Lenders.append("NA in MCA")
        Last_Charge_Activity.append("NA in MCA")
        print("Detail Scraping Progress: "+str(co_num)+" / "+str(Total_Companies)+" Not Available in MCA")
        
    else:
        time.sleep(2.5)
        
        try:
            Co_Name = WebDriverWait(driver, 1, ignored_exceptions=ignored_exceptions).until(
                EC.presence_of_element_located((By.ID, "td_companyname"))
                )
        except:
            pass
        try:
            Company_Legal_Name.append(Co_Name.text)
        except:
            Company_Legal_Name.append("Error due to stale element")
            print("Stale element Error in Company_Legal_Name")
            
        try:
            www = WebDriverWait(driver, 1, ignored_exceptions=ignored_exceptions).until(
                EC.presence_of_element_located((By.ID, "dv_website"))
                )
        except:
            pass
        try:
            Website_1.append("www."+www.text)
        except:
            Website_1.append("Error due to stale element")
            print("Stale element Error in Website_1")
            
        try:
            co_type = WebDriverWait(driver, 1, ignored_exceptions=ignored_exceptions).until(
                EC.presence_of_element_located((By.ID, "td_companyclass"))
                )
        except:
            pass
        try:    
            Company_Type.append(co_type.text)  
        except:
            Company_Type.append("Error due to stale element")
            print("Stale element Error in Company_Type")
        
        try:    
            authcap = WebDriverWait(driver, 1, ignored_exceptions=ignored_exceptions).until(
                EC.presence_of_element_located((By.ID, "td_authorisedCapital"))
                )
        except:
            pass
        try:
            Authorised_Capital.append(authcap.text)
        except:
            Authorised_Capital.append("Error due to stale element")
            print("Stale element Error in Authorised_Capital")
       
        try:    
            paidcap = WebDriverWait(driver, 1, ignored_exceptions=ignored_exceptions).until(
                EC.presence_of_element_located((By.ID, "td_paidupCapital"))
                )
        except:
            pass
        try:
            Paid_up_Capital.append(paidcap.text)
        except:
            Paid_up_Capital.append("Error due to stale element")
            print("Stale element Error in Paid_up_Capital")
      
        try:    
            DAGM = WebDriverWait(driver, 1, ignored_exceptions=ignored_exceptions).until(
                EC.presence_of_element_located((By.ID, "td_dateOfAgm"))
                )
        except:
            pass
        try:
            Date_of_AGM.append(DAGM.text)
        except:
            Date_of_AGM.append("Error due to stale element")
            print("Stale element Error in Date_of_AGM")
        
        try:
            DBS = WebDriverWait(driver, 1, ignored_exceptions=ignored_exceptions).until(
                EC.presence_of_element_located((By.ID, "td_DateOfBalanceSheet"))
                )
        except:
            pass
        try:
            Date_of_Balance_Sheet.append(DBS.text)
        except:
            Date_of_Balance_Sheet.append("Error due to stale element")
            print("Stale element Error in Date_of_Balance_Sheet")
        
        try:
            info_ls = driver.find_elements_by_css_selector("td[id='td_listingStatus']")
        except:
            pass
        info = []
        for i in info_ls:
            try:
                info.append(i.text)
            except:
                info.append("Error due to Stale Element")        
        try:
            Industry.append(info[1])
            Segment.append(info[2])
        except:
            Industry.append("NA")
            Segment.append("NA")
        
        try:
            KP = WebDriverWait(driver, 1, ignored_exceptions=ignored_exceptions).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "b[class='convert d-block border-effect']"))
                )
        except:
            KP = "Company no longer exsist"
        try:
            Key_Person_Name.append(KP.text)
        except:
            Key_Person_Name.append(KP)
            
        try:    
            charges = driver.find_element_by_css_selector("i[class='fas fa-money-bill-alt mr-1']")
            charges.click()
        except:
            pass
        
        try:
            charges_table = WebDriverWait(driver, 1, ignored_exceptions=ignored_exceptions).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='w-50 right_firstBox res-card ml-2']"))
                )
            time.sleep(0.5)
            ptags = charges_table.find_elements_by_css_selector("p[class='m-0']")
            time.sleep(0.5)
            txt = []
            for i in ptags:
                txt.append(i.text)
            
            #print(txt)
            Current_Loan_Amount.append(txt[1])
            Loan_Amount_Satisfied.append(txt[3])
            Total_Lenders.append(txt[5])
            Top_Lender_Name.append(txt[7])
            Last_Charge_Activity.append(txt[9])
            print("Detail Scraping Progress: "+str(co_num)+" / "+str(Total_Companies)+" Loan Found")
        
        except:
            Current_Loan_Amount.append("No Loan Taken")
            Loan_Amount_Satisfied.append("No Loan Taken")
            Top_Lender_Name.append("No Loan Taken")
            Total_Lenders.append("No Loan Taken")
            Last_Charge_Activity.append("No Loan Taken")
            print("Detail Scraping Progress: "+str(co_num)+" / "+str(Total_Companies))
            
        try:
            Com_Overview = driver.find_element_by_css_selector("i[class='far fa-building mr-2']")
            Com_Overview.click() 
        except:
            driver.get("https://www.thecompanycheck.com/company/")
            time.sleep(3)
 
            
# Making the Dataframe
Output_DF = pd.DataFrame(
    {'Company_Name': Company_Name,
     'City_Name': City_Name,
     'State': State,
     'Website': Website,
     'Website_1': Website_1,
     'CIN': CIN,
     'Incorpoation_Date': Incorpoation_Date,
     'Company_Type': Company_Type,
     'Authorised_Capital': Authorised_Capital,
     'Paid_up_Capital': Paid_up_Capital,
     'Date_of_AGM': Date_of_AGM,
     'Date_of_Balance_Sheet': Date_of_Balance_Sheet,
     'Industry': Industry,
     'Segment': Segment,
     'Key_Person_Name': Key_Person_Name,
     'Current_Loan_Amount': Current_Loan_Amount,
     'Loan_Amount_Satisfied': Loan_Amount_Satisfied,
     'Top_Lender_Name': Top_Lender_Name,
     'Total_Lenders': Total_Lenders,
     'Last_Charge_Activity': Last_Charge_Activity,
     'Company_Legal_Name': Company_Legal_Name
     })    

time.sleep(1)
print(" ")

# Saving that dataframe in Excel format
print("Saving scraped data in Excel format in your folder where you have saved this .py file")
datatoexcel = pd.ExcelWriter('Nasscom_Final_Data_Scraped.xlsx')    
Output_DF.to_excel(datatoexcel)
datatoexcel.save()

time.sleep(2)
driver.close()
