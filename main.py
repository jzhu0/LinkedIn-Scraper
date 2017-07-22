'''
Created on Jul 21, 2017


@author: Jason Zhu
'''
DATASET_NAME = "COPY - Founder backgrounds & Operational - 7.5.2017.xlsx"
DATASET_SHEETNAME = "Sheet1"
DOWNLOAD_DIRECTORY = ""

import time
from bs4 import BeautifulSoup
import requests
import openpyxl
from selenium import webdriver

# Extraction functions
def download_profile(driver):
    driver.find_element_by_class_name("action-name-btn").click()

print "Loading data set " + DATASET_NAME + "..."
BgOpData = openpyxl.load_workbook(DATASET_NAME)
links = {}
if (DATASET_SHEETNAME in BgOpData.sheetnames):
    BgOpDataSheet = BgOpData[DATASET_SHEETNAME]
else:
    raise Exception("Data set did not contain correct data sheet")
print "    Data set successfully opened"
for row in BgOpDataSheet.iter_rows():
    if (row[0].value == "Project ID"):   # skip the first row
        continue
    if not (row[2].value == "None"):   # skip rows without a link
        links[row[1].value] = str(row[2].value)
        print "Loaded pID #" + str(row[0].value) + ": " + str(row[1].value)
    if (row[0].value > 10):   # test loading a few rows
        break
print "    Data set successfully loaded"

# test with smaller links dict:
print "Using hard-coded data set."
links = {"Raanan Zehavi" : "https://www.linkedin.com/in/raanan-zehavi-418933b/",
"Aksel Chernitzky" : "https://www.linkedin.com/in/aksel-chernitzky-5902a/",
"Janice Chernitzky" : "https://www.linkedin.com/in/janice-chernitzky-7259b763/",
"Daniel Ferrazzoli" : "https://www.linkedin.com/in/daniel-ferrazzoli-2a76a72a/"}


# Set download location preferences
preferences = {"download.default_directory" : DOWNLOAD_DIRECTORY}
options = webdriver.ChromeOptions()
options.add_experimental_option("prefs", preferences)
driver = webdriver.Chrome(chrome_options = options)

for name, pURL in links.iteritems():
    print "Processing " + name
    req = requests.get(pURL)
    data = BeautifulSoup(req.content, "html.parser")
    time.sleep(0.5)
    
    driver.get(pURL)
    download_profile(driver)
    time.sleep(0.5)

