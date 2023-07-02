# https://sellercentral.amazon.com/revcal?ref=RC1&
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager as CM
import time
from openpyxl import Workbook, load_workbook
# If you need to get the column letter, also import this
from openpyxl.utils import get_column_letter
import requests
import os
import warnings
import argparse
import json
import xlwings as xw

def getProfiles():
	file = open("chrome_profiles.json", "r")
	config = json.load(file)
	return config

def parse(fileinput, profile, country, isreplace, xlsheet):
    warnings.filterwarnings("ignore", category=UserWarning) 
    options = webdriver.ChromeOptions()
    options.add_argument("user-data-dir={}".format(getProfiles()[profile]['chrome_user_data']))
    options.add_argument("profile-directory={}".format(getProfiles()[profile]['chrome_profile']))
    options.add_argument('--no-sandbox')
    options.add_argument("--log-level=3")
    options.add_argument("--window-size=800,600")
    # options.add_argument("user-agent=" + ua.random )
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    driver = webdriver.Chrome(service=Service(CM().install()), options=options)

    driver.get('https://sellercentral.amazon.com/revcal?ref=RC1&')
    cookies = {}
    for cookie in driver.get_cookies():
        cookies[cookie['name']] = cookie['value']
    # print(cookies)
    # exit()
    os.system('cls')
    print('File Selected:', fileinput)
    statistic_source = fileinput
    wb = load_workbook(filename=statistic_source)
    ws = wb['Sheet1']
    # Use the active cell when the file was loaded
    ws = wb.active
    for i in range(1, ws.max_row + 1):
        if ws['A{}'.format(i)].value == None:
            break
        print(ws['A{}'.format(i)].value)

        headers = {
            'authority': 'sellercentral.amazon.com',
            'accept': '*/*',
            'accept-language': 'en-US,en;q=0.9',
            'referer': 'https://sellercentral.amazon.com/revcal?ref=RC1&',
            'sec-ch-ua': '".Not/A)Brand";v="99", "Google Chrome";v="103", "Chromium";v="103"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'sec-fetch-dest': 'empty',
            'sec-fetch-mode': 'cors',
            'sec-fetch-site': 'same-origin',
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36',
        }

        params = {
            'searchKey': '{}'.format(ws['A{}'.format(i)].value),
            'countryCode': '{}'.format(country),
            'locale': 'en-US',
        }
        session = requests.Session()
        response = session.get('https://sellercentral.amazon.com/revenuecalculator/productmatch', params=params, cookies=cookies, headers=headers)
        data = response.json()
        csrftoken = response.headers['anti-csrftoken-a2z']
        try:
            name = data['data']['otherProducts']['products'][0]['title']
            weight = data['data']['otherProducts']['products'][0]['weight']
            dim1 = data['data']['otherProducts']['products'][0]['length']	
            dim2 = data['data']['otherProducts']['products'][0]['width']	
            dim3 = data['data']['otherProducts']['products'][0]['height']	
            asin = data['data']['otherProducts']['products'][0]['asin']	
            try:
                merchantsku = data['data']['otherProducts']['products'][0]['merchantSku']	
            except:
                merchantsku = ''
            try:
                fnsku = data['data']['otherProducts']['products'][0]['fnsku']	
            except:
                fnsku = ''
            gl = data['data']['otherProducts']['products'][0]['gl']
            try:	
                specialdel = data['data']['otherProducts']['products'][0]['specialDeliveryRequirement']	
            except:
                specialdel = ''
            dimensionunit = data['data']['otherProducts']['products'][0]['dimensionUnit']	
            weightunit = data['data']['otherProducts']['products'][0]['weightUnit']
        except:
                try:
                    name = data['data']['myProducts']['products'][0]['title']
                    weight = data['data']['myProducts']['products'][0]['weight']
                    dim1 = data['data']['myProducts']['products'][0]['length']	
                    dim2 = data['data']['myProducts']['products'][0]['width']	
                    dim3 = data['data']['myProducts']['products'][0]['height']	
                    asin = data['data']['myProducts']['products'][0]['asin']
                    try:
                        merchantsku = data['data']['myProducts']['products'][0]['merchantSku']	
                    except:
                        merchantsku = ''
                    try:
                        fnsku = data['data']['myProducts']['products'][0]['fnsku']	
                    except:
                        fnsku = ''
                    gl = data['data']['myProducts']['products'][0]['gl']
                    try:	
                        specialdel = data['data']['myProducts']['products'][0]['specialDeliveryRequirement']
                    except:
                        specialdel = ''
                    dimensionunit = data['data']['myProducts']['products'][0]['dimensionUnit']	
                    weightunit = data['data']['myProducts']['products'][0]['weightUnit']
                except:
                    print('Product is Not Found')
                    # exit()
                    continue

        headers = {
            'authority': 'sellercentral.amazon.com',
            'accept': '*/*',
            'accept-language': 'en-US,en;q=0.9',
            'referer': 'https://sellercentral.amazon.com/revcal?ref=RC1&',
            'sec-ch-ua': '".Not/A)Brand";v="99", "Google Chrome";v="103", "Chromium";v="103"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'sec-fetch-dest': 'empty',
            'sec-fetch-mode': 'cors',
            'sec-fetch-site': 'same-origin',
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36',
        }
        params = {
            'asin': '{}'.format(ws['A{}'.format(i)].value),
            'countryCode': '{}'.format(country),
            'locale': 'en-US',
            'searchType':"GENERAL",
            'fnsku':"",
        }
        # print(params)
        # time.sleep(5)
        session = requests.Session()
        response = session.get('https://sellercentral.amazon.com/revenuecalculator/getadditionalpronductinfo', params=params, cookies=cookies, headers=headers)
        data = response.json()
        # print(data)
        try:
            afnprice = data['data']['price']['amount']
            mfnprice = data['data']['price']['amount']
        except:
            afnprice = "1"
            mfnprice = "1"
        # try:    
        #     currency = data['data']['price']['currency']
        # except:
        #     currency= 'USD'
        if country == 'CA':
            currency = "CAD"
        else:
            currency = "USD"
        try:
            mfnshipping = data['data']['shipping']['amount']
        except:
            mfnshipping = 0
        
        try:
            # input(data)
            weight = data['data']['weight']
            dim1 = data['data']['width']
            dim2 = data['data']['length']
            dim3 = data['data']['height']

        except Exception as e:
            pass

        params = {
            'locale': 'en-US',
        }

        json_data = {
            'countryCode': '{}'.format(country),
            'itemInfo': {
                'asin': asin,
                # 'merchantSku':  merchantsku,
                # 'fnsku': fnsku,
                'glProductGroupName': gl,
                # 'specialDeliveryRequirement': specialdel,
                'packageLength': '{}'.format(dim2),
                'packageWidth': '{}'.format(dim1),
                'packageHeight': '{}'.format(dim3),
                'dimensionUnit': dimensionunit,
                'packageWeight': '{}'.format(weight),
                'weightUnit': weightunit,
                'afnPriceStr': '{}'.format(afnprice),
                'mfnPriceStr': '{}'.format(mfnprice),
                'mfnShippingPriceStr': '{}'.format(mfnshipping),
                'currency': currency,
                'isNewDefined': False,
            },
            'programIdList': [
                'Core',
                'MFN',
            ],
        }
        # print("xxx")
        # input(json_data)

        if merchantsku != '':
            json_data['itemInfo']['merchantSku'] = merchantsku
        if fnsku != '':
            json_data['itemInfo']['fnsku'] = fnsku
        if specialdel != '':
            json_data['itemInfo']['specialDeliveryRequirement'] = specialdel

        headers = {
            'authority': 'sellercentral.amazon.com',
            'accept': '*/*',
            'accept-language': 'en-US,en;q=0.9',
            'anti-csrftoken-a2z': csrftoken, 
            'content-type': 'application/json; charset=UTF-8',
            'origin': 'https://sellercentral.amazon.com',
            'referer': 'https://sellercentral.amazon.com/revcal?ref=RC1&',
            'sec-ch-ua': '".Not/A)Brand";v="99", "Google Chrome";v="103", "Chromium";v="103"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'sec-fetch-dest': 'empty',
            'sec-fetch-mode': 'cors',
            'sec-fetch-site': 'same-origin',
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36',
        }

        
        if dim1 == 0 or dim2 == 0 or dim3 == 0:
            feeamount = 0    
        else:
            session = requests.Session()
            response = session.post('https://sellercentral.amazon.com/revenuecalculator/getfees', params=params, cookies=cookies, headers=headers, json=json_data)
            # input(response)
            data = response.json()
            # input(data)
            try:
                feeamount = data['data']['programFeeResultMap']['Core']['otherFeeInfoMap']['FulfillmentFee']['total']['amount']
            except:
                feeamount = ""
                # input(data)            

        print(name, weight, dim1, dim2, dim3, feeamount)
        ws['B{}'.format(i)].value = name
        ws['C{}'.format(i)].value = round(weight,3)
        ws['D{}'.format(i)].value = round(dim1,3)
        ws['E{}'.format(i)].value = round(dim2,3)
        ws['F{}'.format(i)].value = round(dim3,3)
        ws['G{}'.format(i)].value = feeamount
        # if i == 10:
        #     break

        time.sleep(2)

        # exit()
    wb.save(statistic_source)
    input('Finished...')


def main():
    parser = argparse.ArgumentParser(description="Statistics")
    parser.add_argument('-i', '--input', type=str,help="File Input")
    parser.add_argument('-d', '--profile', type=str,help="Chrome Profile Selected")
    parser.add_argument('-c', '--country', type=str,help="Country")
    parser.add_argument('-r', '--replace', type=str,help="Replace")

    args = parser.parse_args()

    if args.input[-5:] != '.xlsx':
        print('File input have to XLSX file')
        exit()
    isExist = os.path.exists(args.input)
    if isExist == False :
        print('Please check XLSX file')
        exit()
    
    xlbook = xw.Book(args.xlsinput)
    xlsheet = xlbook.sheets["Sheet1"]
    input("")
    parse(args.input, args.profile, args.country, args.replace, xlsheet)
    
if __name__ == '__main__':
    main()
