import pandas as pd
import re
import csv
import time
import getpass
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
        driver.implicitly_wait(1)
        driver.find_element_by_xpath(xpath)
    except NoSuchElementException:
        return False
    return True

def check_exists_link_text(text):
    try:
        driver.implicitly_wait(1)
        driver.find_element_by_link_text(text)
    except NoSuchElementException:
        return False
    return True

def gen(conference):
    stack = []
    for i in conference:
        if not (re.search(r'\d', i) and not re.search('3D', i)) and not re.search(r"^M{0,3}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})$",i):
            yield i
            continue

def get_conference_names(CSVLocation):
    df = pd.read_csv(CSVLocation + '/conferencenames.csv')
    rawconferences = df['Conference Title'].tolist()
    conferences = []
    conferencenames = {}
    for rawconferencetitle in rawconferences:
        if (isinstance(rawconferencetitle, float)): continue
        conference = rawconferencetitle.split(' ')
        g = gen(conference)
        conference = ' '.join(g)
        conference = conference.split('(')[0].strip()
        if conference.lower() in conferencenames:
            continue
        conferencenames[conference.lower()] = None
        conferences.append(conference)

    return(conferences)

def getDataForConference(conferences, download_dir):
    with open(dir + '/numberofconferences.csv', 'w', newline='', encoding='utf8') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=['Conference Title', 'Number of Papers Published'])
        writer.writeheader()
        for conference in conferences:
            row = {}
            driver.find_element_by_id('advancedSearchInputArea').clear()
            driver.find_element_by_id('advancedSearchInputArea').send_keys('CF=(' + conference + '*) AND FPY=2020')
            driver.find_element_by_xpath("//button[@data-ta='run-search']").click()
            if (check_exists_by_xpath("//div[@class='search-error error-code light-red-bg ng-star-inserted']")):
                row['Number of Papers Published'] = 0
                row['Conference Title'] = conference
                writer.writerow(row)
                print('Writing row ' + str(row) + ' to CSV file')
                continue
            if (check_exists_by_xpath("//button[@id='pendo-close-guide-dc656865']")):
                driver.find_element_by_xpath("//button[@id='pendo-close-guide-dc656865']").click()
            driver.implicitly_wait(10)
            maxresults = int(driver.find_element_by_xpath("//span[@class='brand-blue']").text.replace(',', ''))


            row['Number of Papers Published'] = maxresults
            row['Conference Title'] = conference
            writer.writerow(row)
            print('Writing row ' + str(row) + ' to CSV file')
            driver.back()

dir = input("Please Enter Main Directory: ")

try:
    conferences = get_conference_names(dir)
except FileNotFoundError:
    print('Error: Conference Name file does not exist')
    raise SystemExit


chrome_options = webdriver.ChromeOptions()

chrome_options.add_experimental_option("prefs", {"download.default_directory":  dir})
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
chrome_options.add_argument("--incognito")
chrome_options.add_argument("--start-maximized")

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

getDataForConference(conferences, dir)

driver.quit()