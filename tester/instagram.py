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
import time
from urllib.parse import urlparse

cud = "C:\\Users\\User\\AppData\\Local\\Google\\Chrome\\User Data"
cp = "Profile 1"
fp_class = '_aano'
fpd_class = 'x1rg5ohu'
urls = ["https://www.instagram.com/victoryhomescanada/?hl=en"]
options = webdriver.ChromeOptions()
options.add_argument("user-data-dir={}".format(cud))
options.add_argument("profile-directory={}".format(cp))
options.add_argument('--no-sandbox')
options.add_argument("--log-level=3")
options.add_argument("--window-size=800,600")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)
driver = webdriver.Chrome(service=Service(executable_path=os.path.join(os.getcwd(), "chromedriver", "chromedriver.exe")), options=options)
i = 0
purl = urlparse(urls[0])
username = str(purl.path).replace("/","")

driver.get(f"https://www.instagram.com/{username}")
followers_button = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, f'//a[@href="/{username}/followers/"]'))
)
followers_button.click()
followers_popup = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, f'//div[@class="{fp_class}"]'))
)

scroll_script = "arguments[0].scrollTop = arguments[0].scrollHeight;"
# breakpoint()
while True:
    last_count = len(driver.find_elements(By.XPATH, f"//div[@class='{fpd_class}']"))
    driver.execute_script(scroll_script, followers_popup)
    time.sleep(2)  # Add a delay to allow time for the followers to load
    new_count = len(driver.find_elements(By.XPATH, f"//div[@class='{fpd_class}']"))
    if new_count == last_count:
        break  
# fBody  = driver.find_element(By.CSS_SELECTOR, "div._aano")
# //div[@class="_aano"]//li
input("done")

# time.sleep(2)
# followerbutton = driver.find_element(By.CSS_SELECTOR, "a[href='/victoryhomescanada/followers/?hl=en']")
# followerbutton.click()
# time.sleep(2)
# fBody  = driver.find_element(By.CSS_SELECTOR, "div._aano")
# while True:
#     driver.execute_script('arguments[0].scrollTop = arguments[0].scrollTop + arguments[0].offsetHeight;', fBody)
#     time.sleep(2)



# followers_popup = driver.find_element(By.XPATH, '//div[@class="_aano"]')

# input("")