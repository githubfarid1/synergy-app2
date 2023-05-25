from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager as CM
from selenium.webdriver.common.action_chains import ActionChains
import json
import warnings
from urllib.parse import urlencode, urlparse
import time
from random import randint
import pyautogui as pg
import undetected_chromedriver as uc 
import os
import shutil
import xlwings as xw
import argparse
import sys
from sys import platform

def clear_screen():
    try:
        if platform == "win32":
            os.system("cls")
        else:    
            os.system("clear")
    except Exception as er:
        print(er, "Command is not supported")

def getConfig():
	file = open("setting.json", "r")
	config = json.load(file)
	return config

def getProfiles():
	file = open("chrome_profiles.json", "r")
	config = json.load(file)
	return config


def browser_init(userdata):
    warnings.filterwarnings("ignore", category=UserWarning)
    options = webdriver.ChromeOptions()
    options.add_argument("user-data-dir={}".format(getProfiles()[userdata]['chrome_user_data']))
    options.add_argument("profile-directory={}".format(getProfiles()[userdata]['chrome_profile']))
    options.add_argument('--no-sandbox')
    options.add_argument("--log-level=3")
    # options.add_argument("--window-size=1200, 900")
    options.add_argument('--start-maximized')
    options.add_argument("--disable-notifications")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument("--disable-blink-features=AutomationControlled")
    # options.add_experimental_option( "prefs",{'profile.managed_default_content_settings.javascript': 1})
    driver = webdriver.Chrome(service=Service(CM().install()), options=options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})") 
    driver.execute_cdp_cmd("Network.setCacheDisabled", {"cacheDisabled":True})
    return driver


def get_urls(xlsheet, domainwl=[]):
    urlList = []
    maxrow = xlsheet.range('A' + str(xlsheet.cells.last_cell.row)).end('up').row
    for i in range(1, maxrow + 2):
        url = xlsheet[f'A{i}'].value
        domain = urlparse(url).netloc
        # if domain in 'www.walmart.com' or domain == 'www.walmart.ca':
        if domain in domainwl:
            tpl = (url, i)
            urlList.append(tpl)
    return urlList

def walmart_scraper(xlsheet, profilename):
    user_data = profilename
    urlList = get_urls(xlsheet, domainwl=['www.walmart.com','www.walmart.ca'])
    i = 0
    maxrec = len(urlList)
    driver = browser_init(userdata=user_data)
    clear_screen()
    while True:
        if i == maxrec:
            break
        url = urlList[i][0]
        rownum = urlList[i][1]
        print(url, end=" ", flush=True)
        driver.get(url)
        try:
            driver.find_element(By.CSS_SELECTOR, "div#topmessage").text
            print("Failed")
            del driver
            waiting = 120
            print(f'The script was detected as bot, please wait for {waiting} seconds', end=" ", flush=True)
            time.sleep(waiting)
            isExist = os.path.exists(user_data)
            print(isExist)
            if isExist:
                shutil.rmtree(user_data)
            print('OK')
            driver = browser_init(userdata=user_data)
            continue
        except:
            
            print('OK')
            pass

        try:
            title = driver.find_element(By.CSS_SELECTOR, "h1[data-automation='product-title']").text
        except:
                title = ''
        try:
            price = driver.find_element(By.CSS_SELECTOR, "span[data-automation='buybox-price']").text
        except:
            price = ''
        try:
            sale = driver.find_element(By.CSS_SELECTOR, "div[data-automation='mix-match-badge'] span").text
        except:
            sale = ''
        
        print(title, price, sale)
        
        xlsheet[f'B{rownum}'].value = price
        xlsheet[f'C{rownum}'].value = sale
        i += 1     


def walmart_scraper2(xlsheet, profilename):
    user_data = profilename
    urlList = get_urls(xlsheet, domainwl=['www.walmart.com','www.walmart.ca'])
    i = 0
    maxrec = len(urlList)
    driver = browser_init(userdata=user_data)
    while True:
        if i == maxrec:
            break
        url = urlList[i][0]
        rownum = urlList[i][1]
        print(url, end=" ", flush=True)
        driver.get(url)
        try:
            driver.find_element(By.CSS_SELECTOR, "div#topmessage").text
            print("Failed")
            del driver
            waiting = 5
            print(f'The script was detected as bot, please wait for {waiting} seconds', end=" ", flush=True)
            time.sleep(waiting)
            # input("bot detected, press any key...")
            isExist = os.path.exists(getProfiles()[user_data]['chrome_user_data'])
            # print(isExist)
            if isExist:
                shutil.rmtree(getProfiles()[user_data]['chrome_user_data'])
            print('OK')
            driver = browser_init(userdata=user_data)
            time.sleep(1)
            continue
        except:
            print('OK')
            pass

        try:
            title = driver.find_element(By.CSS_SELECTOR, "h1[data-automation='product-title']").text
        except:
                title = ''
        try:
            price = driver.find_element(By.CSS_SELECTOR, "span[data-automation='buybox-price']").text
        except:
            try:
                price = driver.find_element(By.CSS_SELECTOR, "span[itemprop='price']").text
            except:                 
                price = ''


        try:
            sale = driver.find_element(By.CSS_SELECTOR, "div[data-automation='mix-match-badge'] span").text
        except:
            sale = ''
        
        print(title, price, sale)
        
        xlsheet[f'B{rownum}'].value = price
        xlsheet[f'C{rownum}'].value = sale
        i += 1     


def superstore_scraper(xlsheet, profilename):
    user_data = profilename
    urlList = get_urls(xlsheet, domainwl=['www.realcanadiansuperstore.ca'])
    i = 0
    maxrec = len(urlList)
    driver = browser_init(userdata=user_data)
    clear_screen()

    while True:
        if i == maxrec:
            break
        url = urlList[i][0]
        rownum = urlList[i][1]
        print(url, end=" ", flush=True)
        driver.get(url)
        try:
            WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "button[data-track='productAddToCartLocalize'], h1[class='error-page__title']")))
            try:
                mess = driver.find_element(By.CSS_SELECTOR, "h1[class='error-page__title']").text
                print(mess)
            except:
                print('OK')
        except:
            print('Timeout')
  
        try:
            title = driver.find_element(By.CSS_SELECTOR, "h1[class='product-name__item product-name__item--name']").text
        except:
            title = ''

        try:
            price = driver.find_element(By.CSS_SELECTOR, "span[class='price__value selling-price-list__item__price selling-price-list__item__price--now-price__value']").text
        except:
            price = ''
        try:
            sale = driver.find_element(By.CSS_SELECTOR, "del[class='price__value selling-price-list__item__price selling-price-list__item__price--was-price__value']").text
            
        except:
            sale = ''
        
        
        price = price.replace("$", '')
        xlsheet[f'B{rownum}'].value = price
        strsale = ''
        if sale != '':
            strsale = "{} (was {})".format(price, sale)
            xlsheet[f'C{rownum}'].value = strsale
        print(title, price, strsale)
        i += 1
        time.sleep(1)
    
def main():
    parser = argparse.ArgumentParser(description="Amazon Shipment Check")
    parser.add_argument('-xls', '--xlsinput', type=str,help="XLSX File Input")
    parser.add_argument('-sname', '--sheetname', type=str,help="Sheet Name of XLSX file")
    parser.add_argument('-module', '--module', type=str,help="Module Run")
    parser.add_argument('-profile', '--profile', type=str,help="Chrome Profile Name")

    args = parser.parse_args()
    if not (args.xlsinput[-5:] == '.xlsx' or args.xlsinput[-5:] == '.xlsm'):
        input('input the right XLSX or XLSM file')
        sys.exit()
    isExist = os.path.exists(args.xlsinput)
    if not isExist:
        input(args.xlsinput + " does not exist")
        sys.exit()

    if args.module == '':
        input("Module parameter was empty")
        sys.exit()

    xlbook = xw.Book(args.xlsinput)
    xlsheet = xlbook.sheets[args.sheetname]
    if args.module == 'superstore':
        superstore_scraper(xlsheet=xlsheet, profilename=args.profile)
    else:
        walmart_scraper2(xlsheet=xlsheet, profilename=args.profile)

    print("Saving to", args.xlsinput, end=".. ", flush=True)
    xlbook.save(args.xlsinput)
    time.sleep(1)    
    print("OK")
    input("End Process..")    


if __name__ == '__main__':
    main()
