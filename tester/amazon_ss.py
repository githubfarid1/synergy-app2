import os
import argparse
import sys
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
from PyPDF2 import PdfMerger, PdfReader, PdfWriter, generic
import time
import glob
from pylovepdf import ILovePdf
ilovepdf_public_key = "project_public_07fb2f104eed13a200b081a9aa6c3e9e_iB33k4a15e8ff325cc90217ab98feb961721d"

cud = "C:\\Users\\User\\AppData\\Local\\Google\\Chrome\\User Data8"
cp = "Default"
item = "tes"
filepath = r"C:\synergy-data-tester\ss"
ob = Screenshot.Screenshot()
options = webdriver.ChromeOptions()
options.add_argument("user-data-dir={}".format(cud))
options.add_argument("profile-directory={}".format(cp))
options.add_argument('--no-sandbox')
options.add_argument("--log-level=3")
options.add_argument("--window-size=800,600")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)
driver = webdriver.Chrome(service=Service(executable_path=os.path.join(os.getcwd(), "chromedriver", "chromedriver.exe")), options=options)
driver.maximize_window()
url = "https://www.amazon.com/dp/{}".format("B0765Z7GCZ")
driver.get(url)
filename = 'file1.png'
# pdf = fitz.open(os.path.join(filepath, "{}_{}.pdf".format(item,"tmp")))
element = driver.find_element(By.XPATH, '//*[@id="ppd"]')
filepathsave = ob.get_element(driver, element, save_path=r"".join(filepath),image_name=filename)
opdf = PdfWriter()
page = opdf.add_blank_page(width=100, height=200)
left = 100 
bottom = 100 
right = 200 
top = 200 
rect = generic.RectangleObject([left, bottom, right, top])
page.artBox = rect

with open("file1.pdf", "wb") as fp:
    opdf.write(fp)
# pdf.save(os.path.join(filepath, "{}.pdf".format(item)))
input("")