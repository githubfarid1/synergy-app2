from playwright.sync_api import sync_playwright
import random

urls = [
 'https://www.walmart.ca/en/ip/WHOPPERS-Malted-Milk-Candy/6000056442004',
 'https://www.walmart.ca/en/ip/Red-Rose-Orange-Pekoe-Black-Tea/6000196553629',
 'https://www.walmart.ca/en/ip/Maynards-Sour-Cherry-Blasters-Candy-355g/6000153706823',
 'https://www.walmart.ca/en/ip/Arnott-s-Tim-Tam-Original-Chocolate-Cookies-200g/6000188762524',
 'https://www.walmart.ca/en/ip/hersheys-cookies-n-crme-mix-sweet-and-salty-snack/6000196715212',
 'https://www.walmart.ca/en/ip/Wagon-Wheels-Cookies-Dare/6000200896790',
 'https://www.walmart.ca/en/ip/Kraft-Cheez-Whiz-Cheese-Spread/6000153706704',
 'https://www.walmart.ca/en/ip/Peek-Freans-Assorted-Cr-me-Biscuit/6000139306160',
 'https://www.walmart.ca/en/ip/OH-HENRY-Chocolatey-Candy-Bites/6000055334062',
 'https://www.walmart.ca/en/ip/tetley-orange-pekoe-tea/6000109757300',
 'https://www.walmart.ca/en/ip/tim-tam-dark-chocolate/6000198846314#find-in-store-section',
 'https://www.walmart.ca/en/ip/NESCAF-Rich-French-Vanilla-Instant-Coffee-100-g/6000195749780',
 'https://www.walmart.ca/en/ip/swiss-chalet-dipping-sauce-mix/6000098328176',
 'https://www.walmart.ca/en/ip/NESCAF-Rich-French-Vanilla-Instant-Coffee-100-g/6000195749780',
 'https://www.walmart.ca/en/ip/Keg-Steakhouse-Bar-Keg-Steak-Seasoning/6000188762503',
 'https://www.walmart.ca/en/ip/Dare-REAL-JUBES-Original-Jujubes-Candy/6000198419621',
 'https://www.walmart.ca/en/ip/Knorr-Bovril-Chicken-Concentrated-Liquid-Stock/6000136119188',
 'https://www.walmart.ca/en/ip/Maynards-Sour-Patch-Kids-Watermelon-Candy-355G/6000197790446',
 'https://www.walmart.ca/en/ip/Mott-s-Clamato-The-Original-Seasoning-Salt-for-Rimming-Glasses/6000197085785',
 'https://www.walmart.ca/en/ip/LOWNEY-CHERRY-BLOSSOM-Candy/6000054728473',
 'https://www.walmart.ca/en/ip/Kraft-Creamy-Cucumber-Dressing/6000101313741', ##
 'https://www.walmart.ca/en/ip/Maynards-Big-Sour-Patch-Kids-Heads-Candy-185G/6000200036332',
 'https://www.walmart.ca/en/ip/Diana-Sauce-Honey-Garlic/6000142774166',
 'https://www.walmart.ca/en/ip/VH-Honey-Garlic-Cooking-Sauce/6000016939137',
 'https://www.walmart.ca/en/ip/Tim-Hortons-Original-Blended-Coffee-Keurig-K-Cup-30ct/6000199413868',
 'https://www.walmart.ca/en/ip/Keg-Steak-Seasoning/6000199313655',
 'https://www.walmart.ca/en/ip/Kerr-s-Clear-Mints/6000196003420',
 'https://www.walmart.ca/en/ip/Red-Rose-K-Cup-Pods-Black-Tea/6000196157417',
 'https://www.walmart.ca/en/ip/Mars-Bar-Caramel-Filled-Chocolate-Candy-Bars-Fun-Size-Peanut-Free-Halloween-25-Count/6000188101512',
 'https://www.walmart.ca/en/ip/hersheys-smores-variety-kit/6000203220475',

]
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



with sync_playwright() as p:
    browser = p.firefox.launch(headless=False, timeout=5000)
    context = browser.new_context(user_agent=random.choice(userAgentStrings))
    page = context.new_page()

    for url in urls:
        page.goto(url)
        if page.title()=='Verify Your Identity':
            print(page.title())
            browser.close()
            browser = p.firefox.launch(headless=False, timeout=5000)
            context = browser.new_context(user_agent=random.choice(userAgentStrings))
            page = context.new_page()
            continue

        price_element = page.locator("span[data-automation='buybox-price']").first
        if price_element.count() > 0:
            # print(price)
            pricetxt = price_element.text_content().replace("$", "")
        else:
            price_element = page.locator("span[itemprop='price']").first
            if price_element.count() > 0:
                pricetxt = price_element.text_content().replace("$", "")
            else:
                pricetxt = "::price unavailable"

        sale_element = page.locator("div[data-automation='mix-match-badge']").first
        if sale_element.count() > 0:
            saletxt = sale_element.text_content()
        else:
            saletxt = "::sale unavailable"
        
        print(page.title(), pricetxt, saletxt)
