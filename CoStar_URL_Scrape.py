# -*- coding: utf-8 -*-
"""
Created on Mon Mar 28 16:17:04 2022

@author: RBurns
"""


import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException, ElementClickInterceptedException, ElementNotInteractableException
from webdriver_manager.chrome import ChromeDriverManager
from tqdm import tqdm
import numpy as np
import time
import os
import ctypes
from ctypes import windll, wintypes
from uuid import UUID
import sys
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup

os.environ['WDM_LOG_LEVEL'] = '0' # Webdriver_Manager set to silence output
os.environ['WDM_PRINT_FIRST_LINE'] = 'False' # Webdriver_Manager set to silence first-line output on start

def launch_webdriver():
    """
    Instantiate the webdriver application with explicitly defined options.
    
    Options set webdriver to maximized window at start, 

    Parameters
    ----------
    webdriver_path : string, optional
        Alternate filepath to directory containing chromedriver.exe.
        The default is None.

    Returns
    -------
    driver : webdriver
        Initialized webdriver for retrieving and manipulating web pages.
    """
    options = ChromeOptions()
    options.add_argument("--log-level=3")
    options.add_argument("start-maximized")
    options.add_experimental_option('excludeSwitches', ['load-extension', 'enable-automation', 'enable-logging'])
    s=ChromeService(ChromeDriverManager().install()) #match verison of chrome
    driver = webdriver.Chrome(service=s, options=options)
    return driver

def scroll(driver, timeout):
    """
    Scroll down webpage with infinite scroll.

    This function is a slightly modified version of the code from
    artjomb at https://gist.github.com/artjomb/07209e859f9bf0206f76.

    Parameters
    ----------
    driver : webdriver Object
        Initialized webdriver for retrieving and manipulating web pages.
    timeout : int
        The number of seconds to put the driver to sleep before continuing
        to scroll. This can be modified depending on connectivity and webpage
        update speeds, so long as the page is given time to load before
        scrolling further.

    Returns
    -------
    None.

    """
    scroll_pause_time = timeout

    # Get scroll height
    last_height = driver.execute_script("return document.body.scrollHeight")

    while True:
        # Scroll down to bottom
        (driver.
         execute_script("window.scrollTo(0, document.body.scrollHeight);"))

        # Wait to load page
        time.sleep(scroll_pause_time)

        # Calculate new scroll height and compare with last scroll height
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            # If heights are the same it will exit the function
            break
        last_height = new_height
    return

driver = launch_webdriver()

driver.get("https://apt22.mapyourshow.com/8_0/explore/exhibitor-gallery.cfm?featured=false")

time.sleep(5)
while True : 
    scroll(driver, 3)
    try:
        loadmorebutton = WebDriverWait(driver,
                                       5).until(EC.visibility_of_element_located((By.CSS_SELECTOR,
                                                                                  'a.btn-secondary')))
        loadmorebutton.click()
    except TimeoutException:
        break
    
#url_list = driver.find_elements(By.CSS_SELECTOR, "div.flex-Thumb.thumb-Border.mr3.pa2")#"section#exhibitor-results > div > table > tbody > tr > td > h3 > a")

#for url in url_list : 
#    print(url.get_attribute("href"))
    
full_page = BeautifulSoup(driver.page_source) 
full_page.prettify()


url_list = full_page.find_all('a')
url_list = [url.get('href') for url in url_list]
correct_url_list = []
for url in url_list:
    if type(url) != str:
        continue 
    if 'exhibitor-details' in url:
        correct_url_list.append(url)
#print(correct_url_list)


correct_url_list = list(set(correct_url_list))


final_category_list = []
final_website_list = []
booth = []
single_company = []
company_list = []
final_booth = []

for url in correct_url_list : 
    driver.get('https://apt22.mapyourshow.com' + url)
    time.sleep(5)
    name = driver.find_elements(By.CSS_SELECTOR, 'h1')
    description = driver.find_elements(By.CSS_SELECTOR, 'p.js-read-more.animated')
    all_booth = driver.find_elements(By.CSS_SELECTOR, 'a#newfloorplanlink')
    for booth in all_booth :
        final_booth.append(booth.text)
    single_company.append(final_booth)
    website = driver.find_elements(By.CSS_SELECTOR, 'ul > li > a[title = "Visit our website"]')

    try : 
        single_company.append(website[0].text)
    except IndexError :
        single_company.append('No Website')
    category = driver.find_elements(By.CSS_SELECTOR, 'div#js-vue-products > div.grid.grid-3-col.grid_centered > div > h2.f3.ma0.tc > a')
    all_category = driver.find_elements(By.CSS_SELECTOR, 'h2 > a')
    
    single_company.append(name[0].text)
    single_company.append(description[0].text)
    
    for category in all_category :
        final_category_list.append(category.text)
    single_company.append(final_category_list)

    company_list.append(single_company)
    single_company= []
    final_category_list = []
    final_booth = []
    
    
    

df = pd.DataFrame(company_list)
writer = pd.ExcelWriter('NAA.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='sheet1', index=False)
writer.save()

 #   all_category = driver.find_elements(By.CSS_SELECTOR, 'h2 > a')
    
    #for category in all_category :
     #   final_category_list.append(category.text)

    
#driver.get('https://apt22.mapyourshow.com' + correct_url_list[0])
#time.sleep(5)
#website = driver.find_elements(By.CSS_SELECTOR, 'ul > li > a[title = "Visit our website"]')
#print(website[0].text)
  
    


#print(name[0].text)
#print(description[3].text)
#print(category[0].text)



#url_test = correct_url_list[0].get('h1')
#print(url_test)

#company = correct_url_list[1]
#test = BeautifulSoup(company)

#testname = test.find_all('h1')
#print(testname)

    

