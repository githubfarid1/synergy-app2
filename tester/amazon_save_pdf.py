from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager as CM
import os
from Screenshot import Screenshot
ob = Screenshot.Screenshot()
options = webdriver.ChromeOptions()
# options.add_argument("--headless")
# options.add_experimental_option('debuggerAddress', 'localhost:9251')
options.add_argument("user-data-dir={}".format(r'C:/Users/User/AppData/Local/Google/Chrome/User Data8'))
options.add_argument("profile-directory={}".format('Default'))
options.add_argument('--no-sandbox')
options.add_argument("--log-level=3")
options.add_argument("--window-size=800,600")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)
# profile = {"plugins.plugins_list": [{"enabled": False, "name": "Chrome PDF Viewer"}], # Disable Chrome's PDF Viewer
#             "download.default_directory": r'C:\synergy-app2\logs', # disable karena kadang gak jalan di PC lain. Jadi downloadnya tetap ke folder download default
#             "download.extensions_to_open": "applications/pdf",
#             "download.prompt_for_download": False,
#             'profile.default_content_setting_values.automatic_downloads': 1,
#             "download.directory_upgrade": True,
#             "plugins.always_open_pdf_externally": True #It will not show PDF directly in chrome                    
#             }
# options.add_experimental_option("prefs", profile)
# driver = webdriver.Chrome(service=Service(CM(version="114.0.5735.90").install()), options=options)
driver = webdriver.Chrome(service=Service(executable_path=os.path.join(os.getcwd(), "chromedriver", "chromedriver.exe")), options=options)
driver.maximize_window()
driver.get("https://www.amazon.com/dp/B076NVVDQZ")
# pdf = driver.execute_cdp_cmd("Page.printToPDF", {
#   "printBackground": False
# })

# import base64

# with open(r"C:\synergy-app2\logs\file2.pdf", "wb") as f:
#   f.write(base64.b64decode(pdf['data']))
# driver.save_screenshot(r"C:\synergy-app2\logs\file2.png")
element = driver.find_element(By.CSS_SELECTOR, "#xppd")
img_url = ob.get_element(driver, element, save_path=r'r"C:\synergy-app2\logs', image_name='paypal.png')
print(img_url)
input("")