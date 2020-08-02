from selenium import webdriver
from pynput.keyboard import Controller
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager
from Digicon.DigiconAuthentication import username, password
import time


class Deactivation:
    def __init__(self):
        # self.browser = webdriver.Chrome(ChromeDriverManager().install())
        self.browser = webdriver.Chrome("C:/Users/Sony/.wdm/drivers/chromedriver/83.0.4103.39/win32/chromedriver.exe")
        self.browser.maximize_window()
        self.keyboard = Controller()

    def login(self):
        self.browser.get('http://103.44.96.51/index.php/')
        time.sleep(1)
        unpath = self.browser.find_element_by_xpath('//*[@id="uname"]')
        unpath.send_keys(username)
        # time.sleep(1)
        pw = self.browser.find_element_by_xpath('//*[@id="password"]')
        pw.send_keys(password)
        loginbtn = self.browser.find_element_by_xpath('//*[@id="proceed"]')
        time.sleep(7)
        loginbtn.click()
        startok = self.browser.find_element_by_xpath('/html/body/div[1]/div[3]/div/button/span')
        time.sleep(2)
        startok.click()
        time.sleep(2)

    def searchbarclick(self, boxno):
        time.sleep(0.5)
        self.browser.find_element_by_xpath('//*[@id="search_stb"]').click()
        time.sleep(0.5)
        self.keyboard.type(boxno)
        self.browser.find_element_by_xpath('//*[@id="SearchSTB"]').click()
        time.sleep(1.5)

    def deactivatebutton(self):
        # Reactivation button click
        try:
            self.browser.find_element_by_xpath('//*[@id="deactive"]').click()
            time.sleep(1)
            # Reactivation confirmation ok prompt message click
            #self.browser.find_element_by_xpath('//*[option[text()="Deactivation Package Change"]]').click()
            s2 = Select(self.browser.find_element_by_id('id_of_element'))
            s2.select_by_value('6')
            time.sleep(1)
            # success/error prompt message click
            self.browser.find_element_by_xpath('/html/body/div[4]/div[3]/div/button[1]/span')
            time.sleep(0.5)
            check = self.browser.find_element_by_xpath('//*[@id="alertify-ok"]')
            check.click()
            time.sleep(2)
            doublecheck = self.browser.find_element_by_xpath('//*[@id="alertify-ok"]');
            doublecheck.click()
            print('Deactivated')
        except NoSuchElementException:
            print('Already Deactivated')
            return "N"
        return "Y"

    def deactivatebutton2(self):
        try:
            # Reactivation button click
            self.browser.find_element_by_xpath('//*[@id="deactive"]').click()
            time.sleep(1)
            # Reactivation confirmation ok prompt message click
            # self.browser.find_element_by_xpath('//*[option[text()="Deactivation Package Change"]]').click()
            s1 = Select(self.browser.find_element_by_id('deactivation_reason'))
            # select by visible text
            # s1.select_by_visible_text('Deactivation Package Change')
            # select by value
            s1.select_by_value('6')
            time.sleep(1)
            # success/error prompt message click
            self.browser.find_element_by_xpath('/html/body/div[4]/div[3]/div/button[1]/span').click()
            time.sleep(0.5)
            self.browser.find_element_by_xpath('//*[@id="alertify-ok"]').click()
            time.sleep(5)
            print("deactivated confirmed")
            check = self.browser.find_element_by_xpath('//*[@id="alertify-ok"]')
            check.click()
            print('Deactivated')
        except NoSuchElementException:
            print('Already Deactivated')
            return "N"
        return "Y"
