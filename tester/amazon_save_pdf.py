from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager as CM
import os
options = webdriver.ChromeOptions()
# options.add_argument("--headless")
# options.add_experimental_option('debuggerAddress', 'localhost:9251')
# options.add_argument("user-data-dir={}".format(getProfiles()[profile]['chrome_user_data']))
# options.add_argument("profile-directory={}".format(getProfiles()[profile]['chrome_profile']))
options.add_argument('--no-sandbox')
options.add_argument("--log-level=3")
options.add_argument("--window-size=800,600")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)
# driver = webdriver.Chrome(service=Service(CM(version="114.0.5735.90").install()), options=options)
driver = webdriver.Chrome(service=Service(executable_path=os.path.join(os.getcwd(), "chromedriver", "chromedriver.exe")), options=options)

driver.get("https://www.amazon.com/dp/B076NVVDQZ")
input("")