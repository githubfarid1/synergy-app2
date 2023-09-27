import os
import argparse
import sys
from datetime import date, datetime, timedelta
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
import fitz
import math
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
import time
import glob
from pylovepdf.ilovepdf import ILovePdf
ilovepdf_public_key = "project_public_07fb2f104eed13a200b081a9aa6c3e9e_iB33k4a15e8ff325cc90217ab98feb961721d"

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

def createpdf(list, filepath):
    for item in list.keys():
        pdf = fitz.open()
        pages = math.ceil(len(list[item])/2)
        for i in range(0, pages):
            pdf.new_page(pno=-1, width=595, height=842)
        pdf.save(os.path.join(filepath, "{}_{}.pdf".format(item,"tmp"))) 
        

def screenshot(list, chrome_data, filepath):
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
    
    for item in list.keys():
        print("Processing box {}...".format(item) , end=" ", flush=True)
        pdf = fitz.open(os.path.join(filepath, "{}_{}.pdf".format(item,"tmp")))
        for idx, value in enumerate(list[item]):
                # print(idx)
                try:
                    url = "https://www.amazon.com/dp/{}".format(value['asin'])
                    driver.get(url)
                    filename = '{}_{}.png'.format(value['box'], str(idx+1))

                    # METHOD 1: screenshoot directly                
                    # filepathsave = os.path.join(filepath, filename)
                    # driver.save_screenshot(filename=filepathsave)

                    # METHOD 2: screenshoot by element                
                    element = driver.find_element(By.XPATH, '//*[@id="ppd"]')
                    filepathsave = ob.get_element(driver, element, save_path=r"".join(filepath),image_name=filename)


                    page = pdf[math.floor(idx/2)]
                    if (idx % 2) == 0:
                        page.insert_image(fitz.Rect(0, 40, 600, 330),filename=filepathsave)
                        page.insert_text((520.2469787597656, 803.38037109375), str(item), color=fitz.utils.getColor("red"))
                        page.insert_text((50, 30), url, color=fitz.utils.getColor("blue"))
                    else:
                        page.insert_image(fitz.Rect(0, 400, 590, 710),filename=filepathsave)
                        page.insert_text((50, 390), url, color=fitz.utils.getColor("blue"))
                except:
                    print(value['asin'], "failed")

        pdf.save(os.path.join(filepath, "{}.pdf".format(item))) 
        pdf.close()
        os.remove(os.path.join(filepath, "{}_{}.pdf".format(item,"tmp")))
        print("OK")
    [os.remove(os.path.join(filepath, file)) for file in os.listdir(filepath) if file.endswith('.png')]


def join_pdfs(filepath):
    merger = PdfMerger()
    print("Merging PDF files in one PDF File..", end=" ", flush=True)
    time.sleep(1)
    fname = "{}.pdf".format("amazon_ss")

    resultfile = os.path.join(filepath, fname) 
    pdffiles = glob.glob(r"".join(filepath) + "\\" + "*.pdf")
    if len(pdffiles) != 0:
        dictfiles = {}
        for pdffile in pdffiles:
            try:
                basefilename = os.path.basename(pdffile)
                dictfiles[int(basefilename.replace(".pdf",""))] = pdffile
            except:
                continue
        sortedfiles = dict(sorted(dictfiles.items()))
        for file in sortedfiles:
            merger.append(sortedfiles[file])
        merger.write(resultfile)
        print("Finished")
        # print("Compressing PDF File..", end=" ", flush=True)
        # ilovepdf = ILovePdf(ilovepdf_public_key, verify_ssl=True)
        # task = ilovepdf.new_task('compress')
        # task.add_file(resultfile)
        # task.set_output_folder(filepath)
        # task.execute()
        # task.download()
        
        # input("Compressed PDF File Download Done")
        # task.delete_current_task()

        return resultfile
    else:
        print("No pdf files was found:", filepath)
        return ""

def pdf_compress(filepath):
        print("Compressing PDF File..", end=" ", flush=True)
        ilovepdf = ILovePdf(ilovepdf_public_key, verify_ssl=True)
        task = ilovepdf.new_task('compress')
        task.set_output_folder('compressed')
        task.execute()
        task.download()
        task.delete_current_task()
        print("Compressed PDF File Download Done")

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
    createpdf(box_grouped, args.pdfoutput)
    screenshot(box_grouped, args.chromedata, args.pdfoutput)
    fileresult = join_pdfs(args.pdfoutput)
    if fileresult:
        pdf_compress(filepath=fileresult)
    
    input("End Process..")    


if __name__ == '__main__':
    main()
