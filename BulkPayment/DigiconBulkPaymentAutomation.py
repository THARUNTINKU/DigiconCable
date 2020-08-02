from selenium import webdriver
from pynput.keyboard import Key, Controller
from webdriver_manager.chrome import ChromeDriverManager
from Digicon.DigiconAuthentication import username, password, paymentpw
from Digicon.BulkPayment.DigiconBulkPaymentExcel import getfilterdpaidlist
import time


class BulkPayment:
    def __init__(self):
        self.browser = webdriver.Chrome(executable_path='D:/Personal/Cable/Digicon/chromedriver.exe')
        self.browser.maximize_window()

    def login(self):
        self.browser.get('http://103.44.96.51/index.php/')
        self.delayonesec()
        unpath = self.browser.find_element_by_xpath('//*[@id="uname"]')
        unpath.send_keys(username)
        # time.sleep(1)
        pw = self.browser.find_element_by_xpath('//*[@id="password"]')
        pw.send_keys(password)
        loginbtn = self.browser.find_element_by_xpath('//*[@id="proceed"]')
        time.sleep(5)
        loginbtn.click()
        time.sleep(2)
        # startok = self.browser.find_element_by_xpath('/html/body/div[1]/div[3]/div/button/span')
        self.delaytwosec()
        # startok.click()
        # self.delaytwosec()

    def delayonesec(self):
        time.sleep(1)

    def delaytwosec(self):
        time.sleep(2)

    def gotobulkpayment(self):
        customers = self.browser.find_element_by_xpath('//*[@id="nav"]/li[3]/a')
        customers.click()
        time.sleep(1)
        bp = self.browser.find_element_by_xpath('//*[@id="nav"]/li[3]/ul/li[2]/a')
        bp.click()
        time.sleep(2)

    def pressfind(self):
        keyboard.press(Key.ctrl)
        keyboard.press('f')
        keyboard.release('f')
        keyboard.release(Key.ctrl)

    def payment(self, paymentpw2):
        pay = self.browser.find_element_by_xpath('//*[@id="pay"]')
        pay.click()
        self.delaytwosec()
        keyboard.press(Key.enter)
        keyboard.release(Key.enter)
        self.delayonesec()
        self.browser.find_element_by_xpath('//*[@id="password"]')
        self.delayonesec()
        keyboard.type(paymentpw2)

    def checkboxselection(self, vcnumber):
        time.sleep(0.1)
        elements = self.browser.find_element_by_xpath('//*[@value="' + vcnumber + '"]')
        elements.find_element_by_xpath('//*[@id="add[]"]')
        elements.click()

    def confirmclick(self):
        self.delayonesec()
        self.browser.find_element_by_xpath('//*[@id="phpbb"]/div[1]/div[3]/div/button[1]/span').click()


# object creation
objbp = BulkPayment()
# Step-1: Login
objbp.login()
# Step-2: Going to bulk payment section
objbp.gotobulkpayment()

keyboard = Controller()
# objbp.pressfind()
# Step-3: Passing vcno to click on checkbox
vcnolist = getfilterdpaidlist()

for vcno in vcnolist:
    time.sleep(0.1)
    objbp.pressfind()
    time.sleep(0.1)
    keyboard.type(vcno)
    objbp.checkboxselection(vcno)
    objbp.pressfind()
    time.sleep(0.1)
    keyboard.press(Key.backspace)
    time.sleep(0.1)
    keyboard.release(Key.backspace)

# Payment Step-1: Pay button and type password
objbp.payment(paymentpw)

# Payment Step-2: Confirm
# objbp.confirmclick()
