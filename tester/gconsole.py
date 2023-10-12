from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import warnings
from selenium.webdriver.common.action_chains import ActionChains
import os
import time
from selenium.webdriver.common.keys import Keys
# import org.openqa.selenium.Keys;
def browser_init():
    warnings.filterwarnings("ignore", category=UserWarning)
    options = webdriver.ChromeOptions()
    options.add_argument("user-data-dir={}".format("C:\\Users\\User\\AppData\\Local\\Google\\Chrome\\User Data"))
    options.add_argument("profile-directory={}".format("Profile 1"))
    options.add_argument('--no-sandbox')
    options.add_argument("--log-level=3")
    # options.add_argument("--window-size=1200, 900")
    options.add_argument('--start-maximized')
    options.add_argument("--disable-notifications")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    profile = {"plugins.plugins_list": [{"enabled": False, "name": "Chrome PDF Viewer"}], # Disable Chrome's PDF Viewer
                "download.extensions_to_open": "applications/pdf", 
                'profile.default_content_setting_values.automatic_downloads': 1,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "plugins.always_open_pdf_externally": True #It will not show PDF directly in chrome                    
            }
    options.add_experimental_option("prefs", profile)
    return webdriver.Chrome(service=Service(executable_path=os.path.join(os.getcwd(), "chromedriver", "chromedriver.exe")), options=options)


driver = browser_init()
url = "https://search.google.com/search-console?utm_source=about-page&resource_id=sc-domain:snowbirdsweets.ca"
driver.get(url)
# button = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "div[data-text='Performance']")))

time.sleep(1)
driver.find_element(By.CSS_SELECTOR, 'div[data-text="Performance"]').click()
time.sleep(1)
blogurls = ['https://snowbirdsweets.ca/blogs/news/ultimate-ranking-of-canadas-favorite-ketchup-chips', 
            'https://snowbirdsweets.ca/blogs/news/top-10-canadian-exclusive-snacks-2',
            'https://snowbirdsweets.ca/blogs/news/maple-cookies'
]

for blogurl in blogurls:
    # breakpoint()
    driver.find_elements(By.CSS_SELECTOR, 'div.c3pUr > div.OTrxGf > span[class="DPvwYc bquM9e"]')[-1].click()
    time.sleep(1)
    # breakpoint()

    driver.find_elements(By.CSS_SELECTOR, "div#DARUcf")[-1].click()
    time.sleep(1)


    el = driver.find_element(By.CSS_SELECTOR, "input[class='whsOnd zHQkBf']")
    actions = ActionChains(driver)
    actions.send_keys(Keys.CONTROL, "a")
    time.sleep(1)
    actions.move_to_element(el).perform()
    time.sleep(1)
    actions.send_keys(Keys.DELETE)
    time.sleep(1)
    actions.move_to_element(el).perform()
    time.sleep(1)
    actions.send_keys(blogurl)
    time.sleep(1)
    actions.move_to_element(el).perform()
    time.sleep(1)
    driver.find_elements(By.CSS_SELECTOR, 'div[data-id="EBS5u"]')[1].click()    
    time.sleep(3)

    driver.find_elements(By.CSS_SELECTOR, 'div.ak1sAb')[1].find_elements(By.CSS_SELECTOR, 'div.OTrxGf')[1].click()

    time.sleep(1)

    driver.find_element(By.CSS_SELECTOR, 'div[data-value="EuPEfe"]').click()

    time.sleep(1)
    driver.find_elements(By.CSS_SELECTOR, 'div[data-id="EBS5u"]')[-1].click()
    # breakpoint()

    time.sleep(3)
    v1 = driver.find_elements(By.CSS_SELECTOR, 'div[data-column-index="0"]')[-1].find_element(By.CSS_SELECTOR, 'div[class="nnLLaf vtZz6e"]').text
    v2 = driver.find_elements(By.CSS_SELECTOR, 'div[data-column-index="1"]')[-1].find_element(By.CSS_SELECTOR, 'div[class="nnLLaf vtZz6e"]').text
    v3 = driver.find_elements(By.CSS_SELECTOR, 'div[jsname="WKVttf"]')[-1].find_element(By.CSS_SELECTOR, 'span.UwdJ1c').text.split('of')[-1].strip()
    print(v1, v2, v3)
input("")

