import os
import argparse
import sys
from datetime import date, datetime, timedelta
# import amazon_lib as lib
import logging
import xlwings as xw
from pathlib import Path
from Screenshot import Screenshot
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager as CM
import json

logger = logging.getLogger()
logger.setLevel(logging.NOTSET)
logger2 = logging.getLogger()
logger2.setLevel(logging.NOTSET)
def getProfiles():
	file = open("chrome_profiles.json", "r")
	config = json.load(file)
	return config

def data_generator(xlsheet):
    maxrow = xlsheet.range('C' + str(xlsheet.cells.last_cell.row)).end('up').row
    listProduct = []
    for i in range(10, maxrow + 2):
        if str(xlsheet['L{}'.format(i)].value) == 'None':
            continue
        mydict = {"box": round(xlsheet['L{}'.format(i)].value), 'asin':xlsheet['C{}'.format(i)].value}
        listProduct.append(mydict)
    grouped_box = {}
    for p in listProduct:
        box = p["box"]
        if box in grouped_box:
            grouped_box[box].append(p)
        else:
            grouped_box[box] = [p]
    return grouped_box        

def screenshot(list, chrome_data):
    ob = Screenshot.Screenshot()
    options = webdriver.ChromeOptions()
    options.add_argument("user-data-dir={}".format(getProfiles()[chrome_data]['chrome_user_data']))
    options.add_argument("profile-directory={}".format(getProfiles()[chrome_data]['chrome_profile']))
    options.add_argument('--no-sandbox')
    options.add_argument("--log-level=3")
    options.add_argument("--window-size=800,600")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    driver = webdriver.Chrome(service=Service(executable_path=os.path.join(os.getcwd(), "chromedriver", "chromedriver.exe")), options=options)
    driver.maximize_window()

    for item in list:
        for val in item.values():            
            driver.get("https://www.amazon.com/dp/{}".format(val['asin']))
            input("")
def main():
    # clear_screan()
    parser = argparse.ArgumentParser(description="Amazon Shipment")
    parser.add_argument('-xls', '--xlsinput', type=str,help="XLSX File Input")
    parser.add_argument('-sname', '--sheetname', type=str,help="Sheet Name of XLSX file")
    parser.add_argument('-output', '--pdfoutput', type=str,help="PDF output folder")
    parser.add_argument('-cdata', '--chromedata', type=str,help="Chrome User Data Directory")
    args = parser.parse_args()
    if not (args.xlsinput[-5:] == '.xlsx' or args.xlsinput[-5:] == '.xlsm'):
        input('2nd File input have to XLSX or XLSM file')
        sys.exit()
    isExist = os.path.exists(args.xlsinput)
    if not isExist:
        input(args.xlsinput + " does not exist")
        sys.exit()
    isExist = os.path.exists(args.pdfoutput)
    if not isExist:
        input(args.pdfoutput + " folder does not exist")
        sys.exit()

    print('Opening the Source Excel File...', end="", flush=True)
    xlbook = xw.Book(args.xlsinput)
    xlsheet = xlbook.sheets[args.sheetname]
    box_grouped = data_generator(xlsheet=xlsheet)
    input("End Process..")    


if __name__ == '__main__':
    main()
