from playwright.sync_api import sync_playwright



def go_to_google():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()
        page.goto("https://www.google.com")
        page.screenshot(path="google.png")
        print('it is working')
        page.wait_for_timeout(50000)
        browser.close()


if __name__ == '__main__':
    go_to_google()