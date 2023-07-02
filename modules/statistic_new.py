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

def get_xlsdata(xlsheet, isreplace):
    datalist = []
    maxrow = xlsheet.range('A' + str(xlsheet.cells.last_cell.row)).end('up').row
    for i in range(1, maxrow + 2):
        asin = xlsheet[f'A{i}'].value
        tup = (asin, i)
        
        if xlsheet[f'B{i}'].value != None:
            if isreplace:
                datalist.append(tup)
        else:
            datalist.append(tup)

    return datalist

def parse(fileinput, profile, country, datalist, xlsheet):
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
    # os.system('cls')
    print('File Selected:', fileinput)
    for row in range(0, len(datalist)):
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
            'searchKey': '{}'.format(datalist[row][0]),
            'countryCode': '{}'.format(country),
            'locale': 'en-US',
        }
        input(params)

        session = requests.Session()
        response = session.get('https://sellercentral.amazon.com/revenuecalculator/productmatch', params=params, cookies=cookies, headers=headers)
        data = response.json()
        input(data)
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
                    time.sleep(2)
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
            'asin': '{}'.format(datalist[row][0]),
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

        if country == 'CA':
            currency = "CAD"
        else:
            currency = "USD"
        try:
            mfnshipping = data['data']['shipping']['amount']
        except:
            mfnshipping = 0
        
        try:
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
                'glProductGroupName': gl,
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
            data = response.json()
            try:
                feeamount = data['data']['programFeeResultMap']['Core']['otherFeeInfoMap']['FulfillmentFee']['total']['amount']
            except:
                feeamount = ""
        print(name, weight, dim1, dim2, dim3, feeamount)
        xlsheet['B{}'.format(datalist[row][1])].value = name
        xlsheet['C{}'.format(datalist[row][1])].value = round(weight,3)
        xlsheet['D{}'.format(datalist[row][1])].value = round(dim1,3)
        xlsheet['E{}'.format(datalist[row][1])].value = round(dim2,3)
        xlsheet['F{}'.format(datalist[row][1])].value = round(dim3,3)
        xlsheet['G{}'.format(datalist[row][1])].value = feeamount
        # if i == 10:
        #     break

        time.sleep(2)
    input('Finished...')


def main():
    parser = argparse.ArgumentParser(description="Statistics")
    parser.add_argument('-i', '--input', type=str,help="File Input")
    parser.add_argument('-s', '--sheetname', type=str,help="Sheet Name")
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
    
    xlbook = xw.Book(args.input)
    xlsheet = xlbook.sheets[args.sheetname]
    # input("")
    datalist = get_xlsdata(xlsheet=xlsheet, isreplace=args.replace)
    parse(args.input, args.profile, args.country, datalist, xlsheet)
    
if __name__ == '__main__':
    main()

