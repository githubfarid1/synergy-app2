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
            # Requests sorts cookies= alphabetically
            # 'cookie': 'ubid-main=130-3908922-3270858; sid="XetUu4ai7bifPzLfLHP8Vw==|5fnXbjUwIS6qb21zO6OtGu9Mq2TvJMVAoJzoErtnhE8="; __Host-mselc=H4sIAAAAAAAA/6tWSs5MUbJSSsytyjPUS0xOzi/NK9HLT85M0XM0DA0KDncOcgly8bWMUNJRykVSmZtalJyRCFKq52jkZ+pp6GTi6Gvq5OkGUpeNrLAApCQkLMDF29M7wsDFNUipFgAKAR4EdQAAAA==; session-id=146-7378674-3450147; csm-hit=tb:EETFMRQ23JVJNSS7S2P7+s-N98CQT4MKY68RGJB1NT0|1658705269743&t:1658705269743&adb:adblk_no; session-id-time=2289425277l; session-token=bhWDdtJdp8kJ3acOrlXTTmx0x6BJi4nF4gIhHbFPrSdGzyhIDUzVEB7+Z8370q6+KiLzlyf/3HuXp0HfQeo7u2EXCYBWTWNWdqninHDcxmtHSt8m6PvxzCHONCYZ9OnDWlxkSX0GUFM/o18g6ffzb+wviDax+otltFjbhc3pCvrQJhBoow3hbEAtM/eYGycyYKocacymYTG8oKUgjAMipx4KsM4fmIsro0aIbRZtlouidjEw4IBlntFPRYGTId0V; x-main="tGJbBvopA2ojz59XsYYhNNpwvx8Nzgw26NeX?bU?ELJVaamnWgnqyL9PSxOkkzYi"; at-main=Atza|IwEBIM-peKpCY-cfyymPPp2_tya9lTQumvakYjS1ePQC9kQhYYcXkMUyINVQl0-ATUvYjyTnhJh2xApoxiqYg0nQ_EsZqPupogVCQdd8QQM7SB0YNykyuQiTLn1iVNXbPH3DQXtcoMvTF_bKRSXyjpx6_oLbElGmxt7rGzXlctmS_HbiuuLhN9QMJcZbhhHimfTSBp6MxbU5xHg0APVXqwdR1jCC9pFnPoDa7-Lmp2IOalDKFw; sess-at-main="QCWuILPzA3OInwTQ+G1is0Fos2W3CuuxwXOtRVJ6wvc="; sst-main=Sst1|PQGJqoRjyjqPMDWrAM8edRGnCesTf1Zx5A7orf94ozeCdfZJ2gvTP6mNIDX5a5aplT-T6PstKgHHjUejfrJp3mIgHmLvE0mIIaGQQZsM0qKM6OpmmBylbMdYkAITBN9BWObidGxtGu_Aa2kyQ1uA9p02Fm8SMPlKc1us8OnHlnq6Mc-LbXtA7ydXNVs9hHvI4sVjw1ST8xoks4kyxuoq84StFlvhZAuSiCLFCua9lkkRw_2YTgkDHVgY_c6rVeumklTt2f25S0qcPv6uqT6_RKR6-qcbmEhzqHSIR14c5SwVGC0',
            'referer': 'https://sellercentral.amazon.com/revcal?ref=RC1&',
            # 'sec-ch-ua': '"Not.A/Brand";v="99", "Google Chrome";v="103", "Chromium";v="103"',
            'sec-ch-ua': '"Not.A/Brand";v="8", "Chromium";v="114", "Google Chrome";v="114"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'sec-fetch-dest': 'empty',
            'sec-fetch-mode': 'cors',
            'sec-fetch-site': 'same-origin',
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36',
        }
        # headers = {
        #     'authority': 'sellercentral.amazon.com',
        #     'accept': '*/*',
        #     'accept-language': 'en-US,en;q=0.9',
        #     # 'cookie': 'ubid-main=134-2916247-3956466; sid="RKuvscDzYqo2pj6EC6xZSg==|rFEMF0Ui1EtgdZNsSaK0if3w96U0CrXu+BuN5pNtC8w="; __Host-mselc=H4sIAAAAAAAA/6tWSs5MUbJSSsytyjPUS0xOzi/NK9HLT85M0XM0DA0KDncOcgly8bWMUNJRykVSmZtalJyRCFKq52jkZ+pp6GTi6Gvq5OkGUpeNrLAApCQkLMDF29M7wsDFNUipFgAKAR4EdQAAAA==; lc-main=en_US; i18n-prefs=USD; session-id=131-7427739-0234048; sp-cdn="L5Z9:CA"; csm-hit=tb:BCSM1CGNC5V9WNVAWCFE+s-BCSM1CGNC5V9WNVAWCFE|1687961922928&t:1687961922928&adb:adblk_no; session-id-time=2318681929l; session-token=OSAM/q5JNzWa4x6Bs4URSqrtmkMcrO+NDrHx96DKpeXv/gK4zBbozK43JwOCAQYJfym4WKBU7vtj50r3+se0la5ABj7PuzHIIEKNc6kiuebU6ft3lnAg/cBmIkTJ47gYn2ow+nTDIQyW3+A5Juc/iTwuJjCdoKUlY/3t7AkIds7BhHT7jnvO89sqPzcGfkZwQXQl4zT3RhsmSqkXi6Skd1qPEsjXv5VyjEKJHCaXMd2WqvJa5ORiTx74kEdLxMBk; x-main="dTPGR1EUeOl0rlwt6KYxgWgy@y79zqGILIqaEwhJZJFeO8DwqwiEUwoJHIt9cdD1"; at-main=Atza|IwEBIOYwB3Vw6gSC-9mrn55LRLkEb0EM6phy53Pf-yBPU-gIp3NpvkPqV66XMcrULpL-D95n_V0K8y3utdhHU-Kjsycc-hGCVuiOZeV5ITbGjdXoCHzjAj5ykd8_Ydu1XQPf0zaZQoDpCuxljuuNhAKxAdkQxHU_9AC1PwXMKX0IFM1jmUqQhFI8HBd1Pfdrs9mV669wPW8zRwN2ETDIbBOqLQAS_gDLtXke7rpW5F3Bep_DQw; sess-at-main="ZDpcJt3EExmE+bz3zjzLyOFjPxJEGGF6IL0BaaOuPHU="; sst-main=Sst1|PQFl5t7iExTeRMvRjVqBF5t-CSMhh_H1suyQH_6sq0wbPbSp7fL5y2iGCApDjRauZFPSdgzV6epj77x_BnUAfaxnLK6k31ZhTLJbYVTgj-gexZtRFpuqc8XJ_4erx_DODYlWYnyXg_L1G5tVXNu-KXIrPz8m9DstbqGn6soXNdt3A1LxOhi1rpJdYBxgDVRoULtQQQBCw2akA6CcIXo5vLkVO_C1V1cTaxat6lw2rpxLVHIb6oshSnFoIlxiub9IkVkBbQlH1VbGR6RmGidkf9-B42sIQL9vtWI-IiCNVGzdNjE',
        #     'referer': 'https://sellercentral.amazon.com/revcal?ref=RC1',
        #     'sec-ch-ua': '"Not.A/Brand";v="8", "Chromium";v="114", "Google Chrome";v="114"',
        #     'sec-ch-ua-mobile': '?0',
        #     'sec-ch-ua-platform': '"Windows"',
        #     'sec-fetch-dest': 'empty',
        #     'sec-fetch-mode': 'cors',
        #     'sec-fetch-site': 'same-origin',
        #     'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36',
        # }
    
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

