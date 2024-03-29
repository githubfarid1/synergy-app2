from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager as CM
import json
import warnings
from urllib.parse import urlencode, urlparse
import time
import os
import xlwings as xw
import argparse
import sys
from sys import platform
from playwright.sync_api import sync_playwright
import random

userAgentStrings = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 13_4) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/113.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 13.4; rv:109.0) Gecko/20100101 Firefox/113.0",
    "Mozilla/5.0 (X11; Linux i686; rv:109.0) Gecko/20100101 Firefox/113.0",
    "Mozilla/5.0 (X11; Linux x86_64; rv:109.0) Gecko/20100101 Firefox/113.0",
    "Mozilla/5.0 (X11; Ubuntu; Linux i686; rv:109.0) Gecko/20100101 Firefox/113.0",
    "Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:109.0) Gecko/20100101 Firefox/113.0",
    "Mozilla/5.0 (X11; Fedora; Linux x86_64; rv:109.0) Gecko/20100101 Firefox/113.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:102.0) Gecko/20100101 Firefox/102.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 13.4; rv:102.0) Gecko/20100101 Firefox/102.0",
    "Mozilla/5.0 (X11; Linux i686; rv:102.0) Gecko/20100101 Firefox/102.0",
    "Mozilla/5.0 (Linux x86_64; rv:102.0) Gecko/20100101 Firefox/102.0",
    "Mozilla/5.0 (X11; Ubuntu; Linux i686; rv:102.0) Gecko/20100101 Firefox/102.0",
    "Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:102.0) Gecko/20100101 Firefox/102.0",
    "Mozilla/5.0 (X11; Fedora; Linux x86_64; rv:102.0) Gecko/20100101 Firefox/102.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 13_4) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.4 Safari/605.1.15",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36 Edg/113.0.1774.57",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 13_4) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36 Edg/113.0.1774.57"
]

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
    # driver = webdriver.Chrome(service=Service(CM(version="114.0.5735.90").install()), options=options)
    driver = webdriver.Chrome(service=Service(executable_path=os.path.join(os.getcwd(), "chromedriver", "chromedriver.exe")), options=options)

    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})") 
    driver.execute_cdp_cmd("Network.setCacheDisabled", {"cacheDisabled":True})
    return driver

def get_urls(xlsheet, domainwl=[], cost_empty_only=False):
    urlList = []
    maxrow = xlsheet.range('A' + str(xlsheet.cells.last_cell.row)).end('up').row
    for i in range(1, maxrow + 2):
        url = xlsheet[f'A{i}'].value
        domain = urlparse(url).netloc
        if domain in domainwl:
            tpl = (url, i)
            if cost_empty_only == True:
                if xlsheet[f'B{i}'].value == None:
                    urlList.append(tpl)
            else:
                urlList.append(tpl)
    return urlList

def walmart_playwright_scraper(xlsheet, cost_empty_only=False):
    urlList = get_urls(xlsheet, domainwl=['www.walmart.com','www.walmart.ca', "walmart.com", "walmart.ca"],cost_empty_only=cost_empty_only)
    i = 0
    maxrec = len(urlList)
    with sync_playwright() as p:
        browser = p.firefox.launch(headless=True, timeout=10000)
        context = browser.new_context(user_agent=random.choice(userAgentStrings))
        page = context.new_page()
        while True:
            if i == maxrec:
                break
            url = urlList[i][0]
            rownum = urlList[i][1]
            print(url, end=" ", flush=True)
            try:
                page.goto(url)
                if page.title()=='Verify Your Identity' or page.title() == 'Robot or human?':
                    print('Failed')
                    # page.wait_for_timeout(1000)
                    browser.close()
                    del browser
                    browser = p.firefox.launch(headless=True, timeout=10000)
                    context = browser.new_context(user_agent=random.choice(userAgentStrings))
                    page = context.new_page()
                    continue
                else:
                    print('OK')

                price_element = page.locator("div[id='main-buybox'] span[data-automation='buybox-price']").first
                if price_element.count() > 0:
                    # print(price)
                    pricetxt = price_element.text_content().replace("$", "").replace("Now","")
                else:
                    price_element = page.locator("div[data-testid='add-to-cart-section'] span[itemprop='price']").first
                    if price_element.count() > 0:
                        pricetxt = price_element.text_content().replace("$", "").replace("Now","")
                    else:
                        # price_element = page.locator("span[data-automation='buybox-price']").first
                        # if price_element.count() > 0:
                        #     pricetxt = price_element.text_content().replace("$", "").replace("Now","")
                        # else:
                        pricetxt = "None"

                
                sale_element = page.locator("div[id='main-buybox'] div[data-automation='mix-match-badge']").first
                if sale_element.count() > 0:
                    saletxt = sale_element.text_content().replace("View All", "")
                else:
                    saletxt = "None"
                title = page.title()
                print(title, pricetxt, saletxt, end="\n\n")
                xlsheet[f'B{rownum}'].value = ""
                xlsheet[f'C{rownum}'].value = ""
                xlsheet[f'D{rownum}'].value = ""

                xlsheet[f'B{rownum}'].value = pricetxt
                xlsheet[f'C{rownum}'].value = saletxt
                if title == "":
                    xlsheet[f'D{rownum}'].value = "Not Found"

            except Exception as e:
                print('Failed')
                print(e)
                page.wait_for_timeout(2000)
                browser.close()
                del browser
                browser = p.firefox.launch(headless=True, timeout=10000)
                context = browser.new_context(user_agent=random.choice(userAgentStrings))
                page = context.new_page()
                continue

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
    
def superstore_playwright_scraper(xlsheet, cost_empty_only=False):
    notfound = ('404 | Real Canadian Superstore', 'Search Results | Real Canadian Superstore')
    urlList = get_urls(xlsheet, domainwl=['www.realcanadiansuperstore.ca', 'realcanadiansuperstore.ca'],cost_empty_only=cost_empty_only)
    i = 0
    maxrec = len(urlList)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, timeout=10000)
        context = browser.new_context(user_agent=random.choice(userAgentStrings))
        page = context.new_page()
        while True:
            if i == maxrec:
                break
            url = urlList[i][0]
            rownum = urlList[i][1]
            print(url, end=" ", flush=True)
            try:
                page.goto(url)
                page.wait_for_selector("button[data-track='productAddToCartLocalize'], h1[class='error-page__title']")
                title = page.title()
                mess = page.locator("h1[class='error-page__title']")
                if mess.count() > 1:
                    print(mess.text_content())
                else:
                    print('OK')

                price_element = page.locator("div[data-track-product-component='product-details'] span[class='price__value selling-price-list__item__price selling-price-list__item__price--now-price__value']").first
                if price_element.count() > 0:
                    pricetxt = price_element.text_content()
                else:
                    pricetxt = ""

                sale_element = page.locator("div[data-track-product-component='product-details'] del[class='price__value selling-price-list__item__price selling-price-list__item__price--was-price__value']").first
                if sale_element.count() > 0:
                    saletxt = sale_element.text_content()
                else:
                    saletxt = ""

                limit_element = page.locator("div[data-track-product-component='product-details'] p[class='text text--small3 text--left global-color-black product-promo__badge__content']").first
                if limit_element.count() > 0:
                    limittxt = limit_element.text_content()
                else:
                    limittxt = ""

                expires_element = page.locator("div[data-track-product-component='product-details'] p[class='text text--small8 text--left inherit product-promo__message__expiry-date']").first
                if expires_element.count() > 0:
                    expirestxt = expires_element.text_content()
                else:
                    expirestxt = ""

                strprice = pricetxt.replace("$", '')
                
                xlsheet[f'B{rownum}'].value = ""
                xlsheet[f'C{rownum}'].value = ""
                xlsheet[f'D{rownum}'].value = ""
                xlsheet[f'E{rownum}'].value = ""
                xlsheet[f'F{rownum}'].value = ""

                xlsheet[f'B{rownum}'].value = strprice
                strsale = ''
                if saletxt != '':
                    strsale = "{} (was {})".format(pricetxt, saletxt)
                    xlsheet[f'C{rownum}'].value = strsale
                expirestxt = expirestxt.replace("Offer expires","").replace(".","")
                xlsheet[f'D{rownum}'].value = limittxt
                xlsheet[f'E{rownum}'].value = expirestxt.replace("Offer expires","")
                if title in notfound:
                    xlsheet[f'F{rownum}'].value = "Not Found"
                    title = "Item Not Found"
                print(title, strprice, strsale, limittxt, expirestxt, end="\n\n")
                i += 1
                page.wait_for_timeout(1000)
            except Exception as e:
                print('Failed')
                print(e)
                page.wait_for_timeout(2000)
                browser.close()
                del browser
                browser = p.chromium.launch(headless=True, timeout=10000)
                context = browser.new_context(user_agent=random.choice(userAgentStrings))
                page = context.new_page()
                continue

def wholesale_playwright_scraper(xlsheet, cost_empty_only=False):
    notfound = ('404 | Wholesale Club', 'Search Results | Wholesale Club')
    urlList = get_urls(xlsheet, domainwl=['www.wholesaleclub.ca', 'wholesaleclub.ca'],cost_empty_only=cost_empty_only)
    i = 0
    maxrec = len(urlList)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, timeout=10000)
        context = browser.new_context(user_agent=random.choice(userAgentStrings))
        page = context.new_page()
        while True:
            if i == maxrec:
                break
            url = urlList[i][0]
            rownum = urlList[i][1]
            print(url, end=" ", flush=True)
            try:
                page.goto(url)
                page.wait_for_selector("button[data-track='productAddToCartLocalize'], h1[class='error-page__title']")
                title = page.title()
                mess = page.locator("h1[class='error-page__title']")
                if mess.count() > 1:
                    print(mess.text_content())
                else:
                    print('OK')

                price_element = page.locator("div[data-track-product-component='product-details'] span[class='price__value selling-price-list__item__price selling-price-list__item__price--now-price__value']").first
                if price_element.count() > 0:
                    pricetxt = price_element.text_content()
                else:
                    pricetxt = ""

                sale_element = page.locator("div[data-track-product-component='product-details'] del[class='price__value selling-price-list__item__price selling-price-list__item__price--was-price__value']").first
                if sale_element.count() > 0:
                    saletxt = sale_element.text_content()
                else:
                    saletxt = ""

                limit_element = page.locator("div[data-track-product-component='product-details'] p[class='text text--small3 text--left global-color-black product-promo__badge__content']").first
                if limit_element.count() > 0:
                    limittxt = limit_element.text_content()
                else:
                    limittxt = ""

                expires_element = page.locator("div[data-track-product-component='product-details'] p[class='text text--small8 text--left inherit product-promo__message__expiry-date']").first
                if expires_element.count() > 0:
                    expirestxt = expires_element.text_content()
                else:
                    expirestxt = ""

                strprice = pricetxt.replace("$", '')
                
                xlsheet[f'B{rownum}'].value = ""
                xlsheet[f'C{rownum}'].value = ""
                xlsheet[f'D{rownum}'].value = ""
                xlsheet[f'E{rownum}'].value = ""
                xlsheet[f'F{rownum}'].value = ""

                xlsheet[f'B{rownum}'].value = strprice
                strsale = ''
                if saletxt != '':
                    strsale = "{} (was {})".format(pricetxt, saletxt)
                    xlsheet[f'C{rownum}'].value = strsale
                expirestxt = expirestxt.replace("Offer expires","").replace(".","")
                xlsheet[f'D{rownum}'].value = limittxt
                xlsheet[f'E{rownum}'].value = expirestxt.replace("Offer expires","")
                if title in notfound:
                    xlsheet[f'F{rownum}'].value = "Not Found"
                    title = "Item Not Found"
                print(title, strprice, strsale, limittxt, expirestxt, end="\n\n")
                i += 1
                page.wait_for_timeout(1000)
            except Exception as e:
                print('Failed')
                print(e)
                page.wait_for_timeout(2000)
                browser.close()
                del browser
                browser = p.chromium.launch(headless=True, timeout=10000)
                context = browser.new_context(user_agent=random.choice(userAgentStrings))
                page = context.new_page()
                continue

def main():
    parser = argparse.ArgumentParser(description="SUperstore, walmart sscraper")
    parser.add_argument('-xls', '--xlsinput', type=str,help="XLSX File Input")
    parser.add_argument('-sname', '--sheetname', type=str,help="Sheet Name of XLSX file")
    parser.add_argument('-module', '--module', type=str,help="Module Run")
    parser.add_argument('-profile', '--profile', type=str,help="Chrome Profile Name")
    parser.add_argument('-isreplace', '--isreplace', type=str,help="is replace the data")


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

    if args.profile == '':
        input("Profile parameter was empty")
        sys.exit()
    if args.isreplace in ["yes", "no"]:
        if args.isreplace == "yes":
            costempty = False
        else:
            costempty = True
    else:    
        input("isreplace parameter was empty")
        sys.exit()

    
    xlbook = xw.Book(args.xlsinput)
    xlsheet = xlbook.sheets[args.sheetname]

    # maxrun = 10
    # for i in range(1, maxrun+1):
    #     if i > 1:
    #         print("Process will be reapeated")
    #     try:    
            # if args.module == 'superstore':
            #     if i == 1:
            #         superstore_playwright_scraper(xlsheet=xlsheet, cost_empty_only=costempty)
            #     else:
            #         superstore_playwright_scraper(xlsheet=xlsheet, cost_empty_only=True)
            #     # superstore_scraper(xlsheet=xlsheet, profilename=args.profile)
            # elif args.module == 'walmart':
            #     if i == 1:
            #         walmart_playwright_scraper(xlsheet=xlsheet, cost_empty_only=costempty)
            #     else:
            #         walmart_playwright_scraper(xlsheet=xlsheet, cost_empty_only=True)
            # elif args.module == 'wholesaleclub':
            #     if i == 1:
            #         wholesale_playwright_scraper(xlsheet=xlsheet, cost_empty_only=costempty)
            #     else:
            #         wholesale_playwright_scraper(xlsheet=xlsheet, cost_empty_only=True)

            # input("End Process..")
        #     break    
        # except Exception as e:
        #     print(e)
        #     if i == maxrun:
        #         input("Execution Limit reached, Please check the script")
        #     time.sleep(10)
        #     continue
            
    if args.module == 'superstore':
        superstore_playwright_scraper(xlsheet=xlsheet, cost_empty_only=costempty)
    elif args.module == 'walmart':
        walmart_playwright_scraper(xlsheet=xlsheet, cost_empty_only=costempty)
    elif args.module == 'wholesaleclub':
        wholesale_playwright_scraper(xlsheet=xlsheet, cost_empty_only=costempty)

    input("End Process..")

if __name__ == '__main__':
    main()
