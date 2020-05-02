'''
This is a script to fetch amazon user's purchase history and load it as CSV file. Automated the process of
1) Start chrome
2) Go to Amazon Website
3) Log In to User account with username and password
4) Go to History Page 
5) Fetch data i.e. Product name, Purchase date, Price, Tax, Shipping Charge, Total Amount and write it as CSV file 
6) Download Invoices
'''


from selenium import webdriver
from selenium.webdriver.chrome.options import Options
# from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.common.action_chains import ActionChains
# from selenium.common.exceptions import StaleElementReferenceException, NoSuchElementException, WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
import os
import time
import csv

class AmazonHistory():
    def __init__(self):
        self.login_url = "https://www.amazon.in/ap/signin?_encoding=UTF8&ignoreAuthState=1&openid.assoc_handle=inflex&openid.claimed_id=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0%2Fidentifier_select&openid.identity=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0%2Fidentifier_select&openid.mode=checkid_setup&openid.ns=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0&openid.ns.pape=http%3A%2F%2Fspecs.openid.net%2Fextensions%2Fpape%2F1.0&openid.pape.max_auth_age=0&openid.return_to=https%3A%2F%2Fwww.amazon.in%2F%3Fref_%3Dnav_signin&switch_account="
        self.username = "yourusername"
        self.cwd = os.getcwd()
        self.start_chrome()
        print(self.cwd)
        self.password = "your password"
        self.output = [["Date", "Amount", "ID", "Product_name", "Product_link", "Price", "Invoice_path"]]

    def start_chrome(self, download_pdf=True, download_prompt=False, headless=False):
        print('Launching Google Chrome')
        self.downloadPath = os.path.join(os.getcwd(), 'Downloads')
        if not os.path.isdir(self.downloadPath):
            os.makedirs(self.downloadPath)
        options = Options()
        if headless:
            options.add_argument('--headless')
            self.enable_download_in_headless_chrome(self.driver, self.downloadPath)
        options.add_argument('--log_level=3')
        # options.add_argument('--no-sandbox')
        # options.add_argument("--remote-debugging-port=9200")
        options.add_experimental_option(
            "prefs", {
                "behavior": "allow",
                "download.prompt_for_download": download_prompt,
                "plugins.always_open_pdf_externally": download_pdf,
                "download.default_directory": self.downloadPath,
                "safebrowsing.enabled": False,
                "safebrowsing.disable_download_protection": True
            }
        )
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        print(self.cwd)
        self.driver = webdriver.Chrome(options=options,
                                       executable_path=os.path.join(self.cwd,
                                                                    # 'Files',
                                                                    'chromedriver.exe'),
                                       service_log_path='log.txt')
        # self.driver.maximize_window()

        return self.driver

    def enable_download_in_headless_chrome(self, driver, download_dir):
        """
        there is currently a "feature" in chrome where
        headless does not allow file download: https://bugs.chromium.org/p/chromium/issues/detail?id=696481
        This method is a hacky work-around until the official chromedriver support for this.
        Requires chrome version 62.0.3196.0 or above.
        """

        # add missing support for chrome "send_command"  to selenium webdriver
        driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')

        params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': download_dir}}
        command_result = driver.execute("send_command", params)
        # print("response from browser:")
        # for key in command_result:
        #     print("result:" + key + ":" + str(command_result[key]))

    def login_to_amazon(self):
        self.driver.get(self.login_url)
        login_id = WebDriverWait(self.driver, 30).until(EC.element_to_be_clickable((By.ID, "ap_email")))
        login_id.clear()
        login_id.send_keys(self.username)
        continue_button = self.driver.find_element_by_id("continue")
        continue_button.click()
        password = WebDriverWait(self.driver, 30).until(EC.element_to_be_clickable((By.ID, "ap_password")))
        password.send_keys(self.password)
        login_button = self.driver.find_element_by_id("signInSubmit")
        login_button.click()
        while r"https://www.amazon.in/?ref_=nav_signin&" not in self.driver.current_url:
           time.sleep(5)

    def user_history(self):
        click1 = self.driver.find_element_by_id('nav-link-accountList')
        click1.click()
        click2 = self.driver.find_element_by_partial_link_text("Your Orders")
        click2.click()
        orders_container = self.driver.find_element_by_id("ordersContainer")
        orders = orders_container.find_elements_by_xpath('.//*[@class="a-box-group a-spacing-base order"]')
        for order in orders:
            table = order.find_element_by_xpath('.//*[@class="a-box a-color-offset-background order-info"]')
            date_minitable = table.find_element_by_xpath('.//*[@class="a-column a-span3"]')
            date = date_minitable.find_element_by_xpath('.//*[@class="a-color-secondary value"]').text
            amount_table = table.find_element_by_xpath('.//*[@class="a-column a-span2"]')
            amount = amount_table.find_element_by_xpath('.//*[@class="a-color-secondary value"]').text
            amount = amount.strip()
            ID_table = table.find_element_by_xpath('.//*[@class="a-fixed-right-grid-col actions a-col-right"]')
            ID = ID_table.find_element_by_xpath('.//*[@class="a-color-secondary value"]').text
            invoice_button = table.find_element_by_partial_link_text('Invoice')
            invoice_button.click()
            invoice = WebDriverWait(self.driver, 30).until(
                EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "Invoice 1")))
            invoice.click()
            path = r'C:\Users\click\PycharmProjects\Scraping\Downloads'
            while not os.path.exists(os.path.join(path, 'Invoice.pdf')):
                time.sleep(5)
            os.rename(os.path.join(path, 'Invoice.pdf'), os.path.join(path, date + '.pdf'))
            entries = order.find_elements_by_xpath(
                './/*[@class="a-fixed-right-grid-inner a-grid-vertical-align a-grid-top"]')
            for entry in entries:
                products = entry.find_elements_by_xpath('.//*[@class="a-fixed-left-grid-col a-col-right"]')
                for product in products:
                    product_name = product.find_element_by_xpath('.//*[@class="a-link-normal"]').text
                    product_link = product.find_element_by_xpath('.//*[@class="a-link-normal"]').get_attribute("href")
                    price = product.find_element_by_xpath('.//*[@class="a-size-small a-color-price"]').text.strip()
                    self.output.append([date, amount, ID, product_name, product_link, price, os.path.join(path, date + '.pdf')])
        print(self.output)
        print(len(self.output))

    def write_csv_file(self):
        with open('Product_History.csv', 'w', newline='') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerows(self.output)


if __name__ == "__main__":
    amazon = AmazonHistory()
    # amazon.start_chrome()
    amazon.login_to_amazon()
    amazon.user_history()
    amazon.write_csv_file()
