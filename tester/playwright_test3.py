from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    browser = p.webkit.launch()
    page = browser.new_page()
    page.goto("https://google.com/")
    page.screenshot(path="example.png")
    browser.close()