#DATA PREPROCESSING
import pandas as pd #pip3 install pandas
import numpy as np #pip3 install numpy
import time
import datetime

#SQL SERVER
import pyodbc
import sqlalchemy
import urllib


#OS
from os import listdir
from os import remove
from os.path import isfile, join, splitext

#SELENIUM
from selenium import webdriver
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains 
from selenium.webdriver.common.keys import Keys

#UPDATE ROUTINE MONITOR
from updateroutinemonitor import update_routine_monitor

geckodriver_path = ''
start_time = time.time()


def delete_all_files_in_folder(folder):
    for f in listdir(folder):
        if isfile(join(folder, f)):
            remove(join(folder, f))


def wait_for_file(folder):    
    counter = 0
    while counter != 1:
        counter = 0
        for f in listdir(folder):
            if isfile(join(folder, f)):
                counter = counter + 1


def df_row_to_sql_values_string(dataset,row,tablename):
    headers = dataset.columns.values.tolist()
    headers_with_brackets = []
    for h in range(len(headers)):
        header_value = headers[h].replace('.',' ')
        header_value = header_value.replace('/',' ')
        header_value = header_value.replace('\\',' ')
        
        headers_with_brackets.append('['+header_value+']')

    valuesList = []
    values_with_brackets = []

    valuesList = dataset.values[row]
    for v in range(len(valuesList)):
        values_with_brackets.append("'"+str(valuesList[v]).replace("'","''")+"'")
    
    headers_string = ','.join(headers_with_brackets)
    
    values_string = ','.join(values_with_brackets)
    
    insert_values_query = """
    INSERT INTO """ + tablename + ' (' + headers_string + ')' + """
    VALUES(""" + values_string + """)
    """
    return insert_values_query


optionsEditted = Options();
optionsEditted.set_preference("browser.download.folderList",2);
optionsEditted.set_preference("browser.download.manager.showWhenStarting", False)


download_path = './Download'
optionsEditted.set_preference("browser.download.dir",download_path)


delete_all_files_in_folder(download_path)


# In[9]:


mime_types = [
    'text/plain', 
    'application/vnd.ms-excel', 
    'text/csv', 
    'application/csv', 
    'text/comma-separated-values', 
    'application/download', 
    'application/octet-stream', 
    'binary/octet-stream', 
    'application/binary', 
    'application/x-unknown'
]
optionsEditted.set_preference("browser.helperApps.neverAsk.saveToDisk", ",".join(mime_types))

firefox_path = ''
binary = FirefoxBinary(firefox_path)

cap = DesiredCapabilities().FIREFOX
cap["marionette"] = True


#Start browser and go to query page
with open('URL.txt','r') as url_file:
    QUERY_URL = url_file.read()

#RESOURCE FILE LOCATION
browser = webdriver.Firefox(capabilities=cap, executable_path=geckodriver_path, firefox_binary=binary, options=optionsEditted)
browser.get(QUERY_URL)

year_selector = browser.find_element_by_id('year')
year_options = year_selector.find_elements_by_tag_name('option')
year_options[-1].click()

try:
    update_year = int(year_options[-1].get_attribute('innerHTML'))

except:
    update_year = int(datetime.datetime.now().year)

print(update_year)


exportBtn = browser.find_element_by_id('exchangeExport')
importBtn = browser.find_element_by_id('exchangeImport')
periodBtn = browser.find_element_by_id('monthly')

searchBtn = browser.find_element_by_xpath("//button[@class='btn btn-primary']")

exportBtn.click()
searchBtn.click()

time.sleep(10)

downloadBtn = browser.find_element_by_xpath("//a[@ng-click='$ctrl.downloadBigSearch()']")
downloadBtn.click()

time.sleep(15)


importBtn.click()
searchBtn.click()

time.sleep(10)

downloadBtn = browser.find_element_by_xpath("//a[@ng-click='$ctrl.downloadBigSearch()']")
downloadBtn.click()

time.sleep(15)


browser.quit()