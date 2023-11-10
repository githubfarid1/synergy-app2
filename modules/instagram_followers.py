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
from urllib.parse import urlparse

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

def parse(xlbook, xlsheet, profile):
    fp_class = '_aano'
    fpdhead_class = 'x1dm5mii'
    fpd1_class = 'x1rg5ohu'
    fpd2_class = 'x193iq5w'
    maxcheck = 10
    driver = browser_init(profilename=profile)    
    maxrow = xlsheet.range('A' + str(xlsheet.cells.last_cell.row)).end('up').row
    for rownum in range(1, maxrow + 1):
        blogurl = xlsheet[f'A{rownum}'].value
        if blogurl == '':
            break

        purl = urlparse(blogurl)
        username = str(purl.path).replace("/","")
        try:
            xlbook.sheets.add(username)
            print("Sheet", username, "Created...")
        except ValueError as V:
            print(V)
        
        ws_active = xlbook.sheets[username]
        ws_active.api.Move(None, After=xlsheet.api)
        ws_active.clear_contents()
        print('Scrape Instagram with acoount', username, '...', end="", flush=True)
        driver.get(f"https://www.instagram.com/{username}")
        followers_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, f'//a[@href="/{username}/followers/"]'))
        )
        followers_button.click()
        followers_popup = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, f'//div[@class="{fp_class}"]'))
        )
        curcheck = 0
        scroll_script = "arguments[0].scrollTop = arguments[0].scrollHeight;"
        while True:
            last_count = len(driver.find_elements(By.CSS_SELECTOR, f"div.{fpdhead_class}"))
            driver.execute_script(scroll_script, followers_popup)
            time.sleep(1)
            new_count = len(driver.find_elements(By.CSS_SELECTOR, f"div.{fpdhead_class}"))
            # print(new_count, last_count)
            if new_count == last_count:
                curcheck += 1
            else:
                curcheck = 0
                
            if curcheck == maxcheck:
                break
        
        data = driver.find_elements(By.CSS_SELECTOR, f"div.{fpdhead_class}")
        for idx, d in enumerate(data):
            try:
                account = d.find_element(By.CSS_SELECTOR, f"div.{fpd1_class}").text
            except:
                account = ""
            try:
                name = d.find_element(By.CSS_SELECTOR, f"span.{fpd2_class}").text
            except:
                name = ""
            ws_active[f'A{idx+1}'].value = account
            ws_active[f'B{idx+1}'].value = name
                
            # print(idx, account, name)
        print("OK")
        
        

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
    print("OK")
    parse(xlbook=xlbook, xlsheet=xlsheet, profile=args.chromedata)
    input("End Process..")    


if __name__ == '__main__':
    main()
