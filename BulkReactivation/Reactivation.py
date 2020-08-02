from selenium import webdriver
from pynput.keyboard import Controller
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from Digicon.DigiconAuthentication import username, password
import time


class Reactivation:
    def __init__(self):
        self.browser = webdriver.Chrome(executable_path='I:/Digicon/Digicon/chromedriver.exe')
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
        time.sleep(0.1)
        self.browser.find_element_by_xpath('//*[@id="search_stb"]').click()
        time.sleep(0.1)
        self.keyboard.type(boxno)
        self.browser.find_element_by_xpath('//*[@id="SearchSTB"]').click()
        # time.sleep(1)

    def reactivatebutton(self):
        # Reactivation button click
        try:
            self.browser.find_element_by_xpath('//*[@id="active"]').click()
            time.sleep(1)
            # Reactivation confirmation ok prompt message click
            self.browser.find_element_by_xpath('//*[@id="alertify-ok"]').click()
            time.sleep(5)
            # success/error prompt message click
            check = self.browser.find_element_by_xpath('//*[@id="alertify-ok"]')
            check.click()
            time.sleep(2)
            # self.browser.find_element_by_xpath('//*[@id="alertify-ok"]').click()
            print('Reactivated')
        except NoSuchElementException:
            # print('Do payment')
            return "N"
        return "Y"
