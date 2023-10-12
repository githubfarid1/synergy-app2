import os
import argparse
import sys
import xlwings as xw
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import warnings
from selenium.webdriver.common.action_chains import ActionChains
import time
from selenium.webdriver.common.keys import Keys
import json

def getProfiles():
	file = open("chrome_profiles.json", "r")
	config = json.load(file)
	return config



def browser_init(profilename):
    warnings.filterwarnings("ignore", category=UserWarning)
    options = webdriver.ChromeOptions()
    options.add_argument("user-data-dir={}".format(getProfiles()[profilename]['chrome_user_data']))
    options.add_argument("profile-directory={}".format(getProfiles()[profilename]['chrome_profile']))
    options.add_argument('--no-sandbox')
    options.add_argument("--log-level=3")
    # options.add_argument("--window-size=1200, 900")
    options.add_argument('--start-maximized')
    options.add_argument("--disable-notifications")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    return webdriver.Chrome(service=Service(executable_path=os.path.join(os.getcwd(), "chromedriver", "chromedriver.exe")), options=options)

def parse(xlsheet, profile):
    maxrow = xlsheet.range('B' + str(xlsheet.cells.last_cell.row)).end('up').row
    for rownum in range(1, maxrow + 2):
        blogurl = xlsheet[f'B{rownum}'].value
        driver = browser_init(profilename=profile)
        url = "https://search.google.com/search-console/performance/search-analytics?resource_id=sc-domain:snowbirdsweets.ca"
        driver.get(url)
        time.sleep(1)

        driver.find_elements(By.CSS_SELECTOR, 'div.c3pUr > div.OTrxGf > span[class="DPvwYc bquM9e"]')[-1].click()
        time.sleep(1)

        driver.find_elements(By.CSS_SELECTOR, "div#DARUcf")[-1].click()
        time.sleep(1)

        el = driver.find_element(By.CSS_SELECTOR, "input[class='whsOnd zHQkBf']")
        actions = ActionChains(driver)
        actions.move_to_element(el)
        actions.send_keys(blogurl)
        actions.perform()
        time.sleep(1)
        actions.perform()
        time.sleep(1)
        driver.find_elements(By.CSS_SELECTOR, 'div[data-id="EBS5u"]')[1].click()    
        time.sleep(3)

        driver.find_elements(By.CSS_SELECTOR, 'div.ak1sAb')[1].find_elements(By.CSS_SELECTOR, 'div.OTrxGf')[1].click()

        time.sleep(1)

        driver.find_element(By.CSS_SELECTOR, 'div[data-value="EuPEfe"]').click()

        time.sleep(1)
        driver.find_elements(By.CSS_SELECTOR, 'div[data-id="EBS5u"]')[-1].click()

        time.sleep(3)
        v1 = driver.find_elements(By.CSS_SELECTOR, 'div[data-column-index="0"]')[-1].find_element(By.CSS_SELECTOR, 'div[class="nnLLaf vtZz6e"]').text
        v2 = driver.find_elements(By.CSS_SELECTOR, 'div[data-column-index="1"]')[-1].find_element(By.CSS_SELECTOR, 'div[class="nnLLaf vtZz6e"]').text
        v3 = driver.find_elements(By.CSS_SELECTOR, 'div[jsname="WKVttf"]')[-1].find_element(By.CSS_SELECTOR, 'span.UwdJ1c').text.split('of')[-1].strip()
        v1 = "0" if v1 == "" else v1
        v2 = "0" if v2 == "" else v2
        v3 = "0" if v1 == "" else v3
        xlsheet[f'D{rownum}'].value = v1        
        xlsheet[f'E{rownum}'].value = v2        
        xlsheet[f'F{rownum}'].value = v3        
        

def main():
    # clear_screan()
    parser = argparse.ArgumentParser(description="Amazon Shipment")
    parser.add_argument('-xls', '--xlsinput', type=str,help="XLSX File Input")
    parser.add_argument('-sname', '--sheetname', type=str,help="Sheet Name of XLSX file")
    parser.add_argument('-cdata', '--chromedata', type=str,help="Chrome User Data Directory")
    args = parser.parse_args()
    if not (args.xlsinput[-5:] == '.xlsx' or args.xlsinput[-5:] == '.xlsm'):
        input('2nd File input have to XLSX or XLSM file')
        sys.exit()
    isExist = os.path.exists(args.xlsinput)
    if not isExist:
        input(args.xlsinput + " does not exist")
        sys.exit()

    print('Opening the Source Excel File...', end="", flush=True)
    xlbook = xw.Book(args.xlsinput)
    xlsheet = xlbook.sheets[args.sheetname]
    parse(xlsheet=xlsheet, profile=args.chromedata)
    input("End Process..")    


if __name__ == '__main__':
    main()
