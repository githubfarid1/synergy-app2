from single_fdaentry import FdaEntry
import argparse
import sys
from sys import platform
import os
import shutil
import time
import fitz
import unicodedata as ud
import uuid
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import warnings
from random import randint
import glob
import string
from datetime import date, datetime
import json
import xlwings as xw
import logging
from pathvalidate import sanitize_filename

POSX1CODE2 = 514.3499755859375
POSX2CODE2 = 594.415771484375

warnings.filterwarnings("ignore", category=Warning)
logger = logging.getLogger()
logger.setLevel(logging.NOTSET)

def getProfiles():
	file = open("chrome_profiles.json", "r")
	config = json.load(file)
	return config

def explicit_wait():
    time.sleep(randint(1, 3))

def getConfig():
	file = open("setting.json", "r")
	config = json.load(file)
	return config

def format_filename(s):
    valid_chars = "-_.() %s%s" % (string.ascii_letters, string.digits)
    filename = ''.join(c for c in s if c in valid_chars)
    filename = filename.replace(' ','_') # I don't like spaces in filenames.
    return filename

def deltree(folder):
    print("Trying removing", folder, "Folder ...", end=" ", flush=True)
    try:
        shutil.rmtree(folder)
    except OSError as e:
        result = "Error: %s : %s" % (folder, e.strerror)
    result =  "Success"
    print(result)

def pdf_rename_individual(pdfoutput_folder, consignee):
    pdffolder = pdfoutput_folder
    list_of_files = glob.glob(os.path.join(pdffolder, "filename*.pdf") )
    if len(list_of_files) == 0:
        return ""
    
    # isExist = os.path.exists(os.path.join(pdffolder, consignee + ".pdf"))
    # if not isExist:
    #     os.rename(list_of_files[0], os.path.join(pdffolder, consignee + ".pdf"))
    # else:
    #     ts = str(time.time())
    #     os.rename(list_of_files[0], os.path.join(pdffolder, consignee + "_" + str(ts) + ".pdf"))

    # print(list_of_files)
    # sys.exit()
    exceptedfiles = []
    for file in list_of_files:
        if file.find("filename") != -1:
            exceptedfiles.append(file)
    if len(exceptedfiles) == 0:
        return ""
    # consname = "_".join(firstitem[0], firstitem[14]) 
    latest_file = max(exceptedfiles, key=os.path.getctime)
    filename = latest_file
    rfilename = os.path.join(pdffolder, sanitize_filename(consignee) + ".pdf")
    isExist = os.path.exists(rfilename)
    
    if isExist:
        ts = str(time.time())
        rfilename = os.path.join(pdffolder, sanitize_filename(consignee) + "_" + str(ts) + ".pdf")

    os.rename(latest_file, rfilename)

    return rfilename

def research_text(pdfpage, text):
    for i in range(0, len(text)+1):
        rect = pdfpage.search_for(text[0:i],flags=(fitz.TEXT_PRESERVE_WHITESPACE))
        if rect == []:
            break
    lastfound = text[0:i-1]
    tail = text.replace(lastfound, "")
    textsearch = lastfound + " " + tail
    return pdfpage.search_for(textsearch,flags=(fitz.TEXT_PRESERVE_WHITESPACE))

def webentry_update_individual(pdffile,  pdffolder, items):
    # print("Update Web Entry Identification Started..")
    time.sleep(1)
    doc = fitz.open(pdffile)
    page = doc[0]
    red = fitz.utils.getColor("red")

    submitter = page.get_text("block", clip=[100.6500015258789, 271.04034423828125, 185.60845947265625, 283.09893798828125]).strip()
    entry_id = page.get_text("block", clip=(152.7100067138672, 202.04034423828125, 230.7493438720703, 214.09893798828125)).strip()
    for item in items:
        searchtext = item[2][:240]
        rects = page.search_for(searchtext, flags=(fitz.TEXT_PRESERVE_WHITESPACE))
        if rects == []:
            rects = research_text(page, searchtext)
        if rects == []:
            input("Item not found, Report to administrator")
            sys.exit()
        pncode2s = page.get_text("blocks", clip=(POSX1CODE2, rects[0][1]-10, POSX2CODE2, rects[0][3]+10))
        xlworksheet['A{}'.format(item[-1])].value = entry_id
        xlworksheet['X{}'.format(item[-1])].value = "'" + pncode2s[0][4].strip()

    doc.close()
    shutil.copy(pdffile, os.path.join(pdffolder, "tmp.pdf"))
    doc = fitz.open(os.path.join(pdffolder, "tmp.pdf"))
    page = doc[0]
    page.insert_text((520.2469787597656, 803.38037109375), item[8], color=red)
    doc.save(pdffile)
    doc.close()    
    print(item[8] + ".pdf" , "Updated")

    time.sleep(1)

def browser_init(profilename, pdfoutput_folder):
    warnings.filterwarnings("ignore", category=UserWarning)
    config = getConfig()
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
    download_dir = pdfoutput_folder
    profile = {"plugins.plugins_list": [{"enabled": False, "name": "Chrome PDF Viewer"}], # Disable Chrome's PDF Viewer
                "download.default_directory": download_dir, 
                "download.extensions_to_open": "applications/pdf", 
                'profile.default_content_setting_values.automatic_downloads': 1,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "plugins.always_open_pdf_externally": True #It will not show PDF directly in chrome                    
            }
    options.add_experimental_option("prefs", profile)
    return webdriver.Chrome(service=Service(executable_path=os.path.join(os.getcwd(), "chromedriver", "chromedriver.exe")), options=options)

def browser_login(driver):
    loginurl = "https://www.access.fda.gov/oaa/logonFlow.htm?execution=e1s1"
    # driver = self.__driver
    driver.get(loginurl)
    WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input[id='understand']")))
    driver.find_element(By.CSS_SELECTOR, "input[id='understand']").click()
    explicit_wait()
    driver.find_element(By.CSS_SELECTOR, "a[id='login']").click()
    explicit_wait()
    driver.find_element(By.CSS_SELECTOR, "a[title='Prior Notice System Interface']").click()
    explicit_wait()
    driver.find_element(By.CSS_SELECTOR, "img[alt='Create New Web Entry']").click()
    explicit_wait()
    return driver

def clear_screan():
    return
    if platform == "win32":
        os.system("cls")
    else:
        os.system("clear")

def clearlist(*args):
    for varlist in args:
        varlist.clear()

def xls_dataframe_generator(filename, sname):
    df = pd.read_excel(filename, sheet_name=sname)
    cols = df.groupby('Shiper Address').first().values.tolist()
    print(cols)

def xls_data_generator(xlws, maxrow):
    global xlworksheet
    global MAXROW
    xlworksheet = xlws #xlworkbook.sheets[sname]
    MAXROW = maxrow
    allData = {}
    wcode = []
    wshipper = []
    wdesc = []
    wsize = []
    wtotal = []
    wmanufact = []
    wmanufact_addr = []
    wmanufact_city = []
    wconsignee = []
    wconsignee_addr = []
    wconsignee_city = []
    wconsignee_postal = []
    wconsignee_state = []
    wconsignee_stact = []
    wsubmitter = []
    wsubmitter_add = []
    wsubmitter_cityetc = []
    wsubmitter_country = []
    wpnumber = []
    wbox = []
    wentrycode = []
    wsku = []
    wrow = []
    wentryid = xlworksheet['B{}'.format(2)].value
    for i in range(2, MAXROW+2):
        if wentryid != xlworksheet['B{}'.format(i)].value:# and xlworksheet['B{}'.format(i)].value != None:
            rid = uuid.uuid4().hex
            allData[rid] = {'data':list(zip(wshipper, wcode, wdesc, wsize, wtotal, wmanufact, wmanufact_addr, wmanufact_city, wconsignee, wconsignee_addr, wconsignee_city, wconsignee_postal, wconsignee_stact, wconsignee_state, wsubmitter, wsubmitter_add, wsubmitter_cityetc, wsubmitter_country, wpnumber, wbox, wentrycode, wsku, wrow)),
            'count' : len(wcode)} 
            wentryid = xlworksheet['B{}'.format(i)].value
            clearlist(wshipper, wcode, wdesc, wsize, wtotal, wmanufact, wmanufact_addr, wmanufact_city, wconsignee, wconsignee_addr, wconsignee_city, wconsignee_postal, wconsignee_stact, wconsignee_state, wsubmitter, wsubmitter_add, wsubmitter_cityetc, wsubmitter_country, wpnumber, wbox, wentrycode, wsku, wrow)
        
        if xlworksheet['B{}'.format(i)].value == None:
            break
        
        wshipper.append(str(xlworksheet['B{}'.format(i)].value).strip())
        wcode.append(str(xlworksheet['F{}'.format(i)].value).strip())
        strdesc= ud.normalize('NFKD', str(xlworksheet['G{}'.format(i)].value).strip()).encode('ascii', 'ignore').decode('ascii')
        wdesc.append(strdesc)
        try:
            wsize.append(str(int(xlworksheet['H{}'.format(i)].options(numbers=int).value)).strip())
        except:
            wsize.append("None")

        wtotal.append(str(xlworksheet['I{}'.format(i)].options(numbers=int).value).strip())
        strmanufact = ud.normalize('NFKD', str(xlworksheet['K{}'.format(i)].value).strip()).encode('ascii', 'ignore').decode('ascii')
        wmanufact.append(strmanufact)
        strmanufact_addr = ud.normalize('NFKD', str(xlworksheet['L{}'.format(i)].value).strip()).encode('ascii', 'ignore').decode('ascii')
        wmanufact_addr.append(strmanufact_addr)
        wmanufact_city.append(str(xlworksheet['M{}'.format(i)].value).strip())
        wconsignee.append(str(xlworksheet['N{}'.format(i)].value).strip())
        wconsignee_addr.append(str(xlworksheet['O{}'.format(i)].value).strip())
        wconsignee_city.append(str(xlworksheet['P{}'.format(i)].value).strip())
        wconsignee_postal.append(str(xlworksheet['Q{}'.format(i)].value).strip())
        wconsignee_state.append(str(xlworksheet['R{}'.format(i)].value).strip())
        wconsignee_stact.append(str(xlworksheet['S{}'.format(i)].value).strip())
        wsubmitter.append(str(xlworksheet['T{}'.format(i)].value).strip())
        wsubmitter_add.append(str(xlworksheet['U{}'.format(i)].value).strip())
        wsubmitter_cityetc.append(str(xlworksheet['V{}'.format(i)].value).strip())
        wsubmitter_country.append(str(xlworksheet['W{}'.format(i)].value).strip())
        wpnumber.append("")
        wbox.append(str(xlworksheet['D{}'.format(i)].value).strip())
        wentrycode.append(str(xlworksheet['A{}'.format(i)].value).strip())
        wsku.append(str(xlworksheet['E{}'.format(i)].value).strip())
        wrow.append(i)
    return allData

def main():
    parser = argparse.ArgumentParser(description="FDA Entry + PDF Extractor")
    parser.add_argument('-i', '--input', type=str,help="File Input")
    parser.add_argument('-s', '--sheet', type=str,help="Sheet Name")
    parser.add_argument('-dt', '--date', type=str,help="Arrival Date")
    parser.add_argument('-d', '--chromedata', type=str,help="Chrome User Data Directory")
    parser.add_argument('-o', '--output', type=str,help="PDF output folder")
    
    args = parser.parse_args()
    if not (args.input[-5:] == '.xlsx' or args.input[-5:] == '.xlsm'):
        input('input the right XLSX or XLSM file')
        sys.exit()
    isExist = os.path.exists(args.input)
    if isExist == False :
        input('Please check XLSX or XLSM file')
        sys.exit()
    if len(args.date) != 10:
        input('Date Arrival is wrong')
        sys.exit()

    isExist = os.path.exists(args.output)
    if isExist == False :
        input('Please make sure PDF folder is exist')
        sys.exit()

    clear_screan()
    file_handler = logging.FileHandler('logs/autofda-err.log')
    file_handler.setLevel(logging.ERROR)
    file_handler_format = '%(asctime)s | %(levelname)s | %(lineno)d: %(message)s'
    file_handler.setFormatter(logging.Formatter(file_handler_format))
    logger.addHandler(file_handler)

    print('Opening the Source Excel File...', end="", flush=True)
    xlbook = xw.Book(args.input)
    print('OK')
    strdate = str(date.today())
    foldernamepn =  os.path.join(args.output, 'prior_notice_{}'.format(strdate))
    isExist = os.path.exists(foldernamepn)
    if not isExist:
        os.makedirs(foldernamepn)
    xlsfilename = os.path.basename(args.input)
    foldername = format_filename("{}_{}_{}".format(xlsfilename.replace(".xlsx", "").replace(".xlsm", ""), args.sheet, strdate) )
    complete_output_folder = os.path.join(foldernamepn, foldername)
    isExist = os.path.exists(complete_output_folder)
    if not isExist:
        os.makedirs(complete_output_folder)
    xlsheet = xlbook.sheets[args.sheet]

    strnow = datetime.now().strftime("%Y-%m-%d-%H%M%S")
    
    filename = "fda-excel-report-{}.log".format(strnow)
    pathname = os.path.join(args.output, filename)

    if os.path.exists(pathname):
        os.remove(pathname)
    file_handler = logging.FileHandler(pathname)
    file_handler.setLevel(logging.CRITICAL)
    
    file_handler_format = '%(message)s'
    file_handler.setFormatter(logging.Formatter(file_handler_format))
    logger.addHandler(file_handler)
    logger.critical("###### Start ######")
    logger.critical("Filename: {}".format(args.input))
    logger.critical("Sheet Name:{}".format(args.sheet))


    maxrow = xlsheet.range('B' + str(xlsheet.cells.last_cell.row)).end('up').row
    xlsdictall = xls_data_generator(xlws=xlsheet, maxrow=maxrow)
   							 							 
    colchecks = {(21, '"SKU"'), (1, '"Code"'), (2, '"Description"'), (3, '"Size (grams)"'), (4, '"Total Quantitiy"'), (5, '"Manufacturer"'), (6, '"Manufacturer address"'), (7, '"Manufacturer City/postal code"'), (8, '"Consignee"'), (9, '"Consignee Address"'), (10, '"Consignee City"'),(11, '"Consignee Postal"'),(12, '"State Actual"'),(13, '"State"'),(14 , '"Shipper/Exporter"'), (15, '"Address"'),(16, '"City/State/Zip Code"'),(17, '"Country"')}
    errors = []
    for idx, xls in xlsdictall.items():
        for data in xls['data']:
            for col in colchecks:
                if data[col[0]] == 'None' or data[col[0]] == '0' or data[col[0]].strip() == '':
                    errors.append((col[1], data[22]))
    
    logger.critical("")
    if len(errors) == 0:
        logger.critical("No Error found in the Excel file")

    else:
        print("Error Found in the excel file. Please check file {}".format(pathname))
        for er in errors:
            logger.critical("Empty or zero value found in column {}, row number {}".format(er[0], er[1]) )

        sys.exit()


    maxrow = xlsheet.range('B' + str(xlsheet.cells.last_cell.row)).end('up').row
    xlsdictall = xls_data_generator(xlws=xlsheet, maxrow=maxrow)
    # input(json.dumps(xlsdictall))
    xlsdictwcode = {}
    for idx, xls in xlsdictall.items():
        for data in xls['data']:
            if data[20] == 'None':
                xlsdictwcode[idx] = xls
                break
    # input(json.dumps(xlsdictwcode))
    # sys.exit()
    for xlsdata in xlsdictwcode.values():
        try:
            driver.close()
            driver.quit()
        except:
            pass
        
        driver = browser_init(profilename=args.chromedata, pdfoutput_folder=complete_output_folder)
        driver = browser_login(driver)

        grouped_cons = {}
        for data in xlsdata['data']:
            cons = (data[8], data[9], data[10], data[11], data[12])
            idx = "#".join(cons) 
            if idx in grouped_cons:
                grouped_cons[idx].append(data)
            else:
                grouped_cons[idx] = [data]
        # print(json.dumps(grouped_cons))
        for item in grouped_cons:
            dl = {}
            dl['data'] = grouped_cons[item]
            dl['count'] = len(dl['data'])
            input(json.dumps(dl))
            fda_entry = FdaEntry(driver=driver, datalist=dl, datearrival=args.date, pdfoutput=complete_output_folder)
            try:
                driver.find_element(By.CSS_SELECTOR, "img[alt='Create WebEntry Button']").click()
            except:
                pass
            fda_entry.parse()
            pdf_filename = pdf_rename_individual(pdfoutput_folder=complete_output_folder, consignee=grouped_cons[item][0][8])
            if pdf_filename != "":
                time.sleep(2)
                webentry_update_individual(pdffile=pdf_filename, pdffolder=complete_output_folder, items=grouped_cons[item])
            else:
                print("file:", pdf_filename)
                input("rename the file was failed")
            print("------------------------------------")
            
    input("data generating completed...")

if __name__ == '__main__':
    main()
