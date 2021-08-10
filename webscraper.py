import selenium
import os
import time
import csv
import pandas as pd
import glob
import getpass
from os import path
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import WebDriverException


main_link = 'http://apps.webofknowledge.com.proxy2.library.illinois.edu/WOS_GeneralSearch_input.do?product=WOS&search_mode=GeneralSearch&SID=8FmeEkA8f3uKTYDYfhK&preferencesSaved='


def login(username, password):
    driver.find_element_by_id('j_username').clear()
    driver.find_element_by_id('j_username').send_keys(username)
    driver.find_element_by_id('j_password').clear()
    driver.find_element_by_id('j_password').send_keys(password)
    driver.find_element_by_xpath("//input[@class='btn btn-primary']").click()
    driver.implicitly_wait(10)

def check_exists_by_xpath(xpath):
    try:
        driver.implicitly_wait(2)
        driver.find_element_by_xpath(xpath)
    except NoSuchElementException:
        return False
    return True

def check_exists_link_text(text):
    try:
        driver.implicitly_wait(2)
        driver.find_element_by_link_text(text)
    except NoSuchElementException:
        return False
    return True

def getDataForCompany(company, download_dir):

    driver.find_element_by_id('advancedSearchInputArea').clear()
    driver.find_element_by_id('advancedSearchInputArea').send_keys('OG=(' + company + '*)')
    driver.find_element_by_xpath("//button[@data-ta='run-search']").click()
    if(check_exists_by_xpath("//div[@class='search-error error-code light-red-bg ng-star-inserted']")):
        print('Search Error for ' + company)
        return False
    if(check_exists_by_xpath("//button[@id='pendo-close-guide-dc656865']")):
        driver.find_element_by_xpath("//button[@id='pendo-close-guide-dc656865']").click()
    maxresults = int(driver.find_element_by_xpath("//span[@class='brand-blue']").text.replace(',', ''))
    markfrom = 1
    while (markfrom <= maxresults):
        num = 1000
        driver.find_element_by_xpath(
            "//button[@class='mat-focus-indicator mat-menu-trigger cdx-but-md cdx-but-white-background margin-right-10--reversible mat-button mat-stroked-button mat-button-base mat-primary']").click()
        driver.find_element_by_xpath("//button[@id='exportToFieldTaggedButton']").click()
        markto = markfrom + num - 1
        if (markto > maxresults):
            markto = maxresults
        driver.find_element_by_xpath(".//mat-radio-button[@value='fromRange']"
                                     "/label[@class='mat-radio-label']"
                                     "/span[@class='mat-radio-container']").click()
        driver.find_element_by_xpath("//input[@aria-label='Input starting record range']").clear()
        driver.find_element_by_xpath("//input[@aria-label='Input starting record range']").send_keys(markfrom)
        driver.find_element_by_xpath("//input[@aria-label='Input ending record range']").clear()
        driver.find_element_by_xpath("//input[@aria-label='Input ending record range']").send_keys(markto)
        driver.find_element_by_xpath("//button[@aria-label=' Author, Title, Source']").click()
        driver.find_element_by_xpath("//div[@aria-label='Full Record']").click()
        driver.find_element_by_xpath("//button[@class='mat-focus-indicator cdx-but-md mat-stroked-button mat-button-base mat-primary']").click()
        time.sleep(1)
        c = 1
        while os.path.exists(download_dir):

            if(c % 10 == 0 and not c == 0):
                driver.find_element_by_xpath(
                    "//button[@class='mat-focus-indicator cdx-but-md mat-stroked-button mat-button-base mat-primary cdk-focused cdk-mouse-focused']").click()

            time.sleep(1)
            if (os.path.isfile(download_dir + '/savedrecs.txt')):
                if(c == 1):
                    print('File exists, renaming file: ' + company + ' file ' + str(markfrom) + ' to ' + str(markto))
                else:
                    print('\nFile exists, renaming file: ' + company + ' file ' + str(markfrom) + ' to ' + str(markto))

                try:
                    os.rename(download_dir + '/savedrecs.txt',
                              download_dir + '//'+ company + ' ' + str(markfrom) + '-' + str(markto) + '.txt')
                except FileExistsError:
                    print('Error: The file: ' + company + ' ' + str(markfrom) + '-' + str(markto) + '.txt' + ' already exists')
                    raise SystemExit
                break
            else:
                if(c == 1):
                    print('Waiting for file to be downloaded', end = '')
                else:
                    print('.', end = '')
            c = c - 1

        markfrom += num

        if (markto == maxresults):
            print('Done downloading ' + company)

def getCSVCompanyContent(CSVLocation):
    df = pd.ExcelFile(CSVLocation).parse('Sheet1')  # you could add index_col=0 if there's an index
    companys = df['Firm Name'].tolist()
    return(companys)



def load_obj(name):
    with open('obj/' + name + '.pkl', 'rb') as f:
        return pickle.load(f)

# Selenium Web Driver Initialization
dir = input("Please Enter Main Directory: ")

try:
    companys = getCSVCompanyContent(dir + '/firmnames.xlsx')
except FileNotFoundError:
    print('Error: Firm Name file does not exist')
    raise SystemExit


chrome_options = webdriver.ChromeOptions()

chrome_options.add_experimental_option("prefs", {"download.default_directory":  dir})
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--incognito")

try:
    driver = webdriver.Chrome(dir + '/chromedriver.exe', options=chrome_options)
except WebDriverException:
    print('Error: Chrome Driver does not exist')
    raise SystemExit
driver.get(main_link)


user = input('Input Net ID Username: ')
pswd = getpass.getpass(prompt='Input Net ID Password: ')

login(user, pswd)

while(True):
    if (check_exists_by_xpath("//div[@class='alert alert-danger']")):
        print("Error: Login Error")
        driver.quit()
        raise SystemExit
    if (check_exists_by_xpath("//button[@id='pendo-close-guide-8fdced48']")):
        driver.find_element_by_xpath("//button[@id='pendo-close-guide-8fdced48']").click()
        if (check_exists_link_text("Advanced Search")):
            driver.find_element_by_link_text("Advanced Search").click()
            break

time.sleep(2)
driver.find_element_by_xpath("//button[@data-ta='add-timespan-row']").click()
driver.find_element_by_xpath("//input[@data-ta='search-timespan-start-input']").send_keys('2011-01-01')
driver.find_element_by_xpath("//input[@data-ta='search-timespan-end-input']").send_keys('2021-07-11')
driver.implicitly_wait(10)

for company in companys:
    if(not getDataForCompany(company, dir) == False):
        driver.back()

driver.quit()







