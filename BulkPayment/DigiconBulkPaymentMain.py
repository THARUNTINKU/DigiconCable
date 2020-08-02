from Digicon.BulkPayment.DigiconBulkPaymentAutomation import BulkPayment
from selenium import webdriver
from pynput.keyboard import Key, Controller
from webdriver_manager.chrome import ChromeDriverManager
from Digicon.DigiconAuthentication import username, password, paymentpw
from Digicon.BulkPayment.DigiconBulkPaymentExcel import getfilterdpaidlist
import time


class DigiconBulkPay:

    def activation():
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
            objbp.pressfind(keyboard)
            time.sleep(0.1)
            keyboard.type(vcno)
            objbp.checkboxselection(vcno)
            objbp.pressfind(keyboard)
            time.sleep(0.1)
            keyboard.press(Key.backspace)
            time.sleep(0.1)
            keyboard.release(Key.backspace)

        # Payment Step-1: Pay button and type password
        # objbp.payment(paymentpw, keyboard)

        # Payment Step-2: Confirm
        # objbp.confirmclick()

    def deactivation():
        objbp = BulkPayment()
        # Step-1: Login
        objbp.login()


objBulkPayment = DigiconBulkPay
objBulkPayment.activation()
# objBulkPayment.deactivation()
