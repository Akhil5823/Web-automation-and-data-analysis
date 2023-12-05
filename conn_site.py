from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains
import time
from selenium.webdriver.firefox.service import Service as FirefoxService
from webdriver_manager.firefox import GeckoDriverManager

from pyvirtualdisplay import Display




def jira():
    # display = Display(visible=0, size=(800, 600))
    # display.start()
    # op=webdriver.ChromeOptions()
    # p={'download.default_directory':r'C:\Users\ebinakh\OneDrive - Ericsson\Documents\Ak_ericsson\tab_to_pow'}
    # op.add_experimental_option('prefs',p)
    file1 = open(r"Kaarunya Files Final 10 Feb\CONFIG_1.txt","r")
    sig=file1.readline()
    pwd=file1.readline()
    driver=webdriver.Firefox()

    # driver.fullscreen_window()

    driver.get("https://eteamproject.internal.ericsson.com/login.jsp") #go to site

    driver.find_element(By.ID,"login-form-username").send_keys(sig)
    driver.find_element(By.ID,"login-form-password").send_keys(pwd)
    driver.find_element(By.ID,"login-form-submit").click()#login

    driver.implicitly_wait(10)
    driver.find_element(By.ID,"find_link").click()

    driver.implicitly_wait(5)

    driver.find_element(By.ID,"filter_lnk_198608_lnk").click()
    driver.find_element(By.ID,"AJS_DROPDOWN__77").click()
    driver.find_element(By.ID,"allExcelFields").click()


def replir():
    driver=webdriver.ChromiumEdge()
    driver.get("https://replir.internal.ericsson.com/authentication/loginsso")
    # driver.get("https://replir.internal.ericsson.com/reports/report-generator")

    # driver.switch_to().alert().accept()

    # driver.switch_to().alert().sendKeys("EBINAKH");
    time.sleep(10)
    
    driver.find_element(By.CLASS_NAME,"ant-btn-primary").click()
    time.sleep(10)
    driver.find_element(By.XPATH,"//span[@class='title item ng-star-inserted'][normalize-space()='Reports']").click()
    time.sleep(2)
    driver.find_element(By.XPATH,"//a[normalize-space()='Report Generator']").click()
    # driver.find_element(By.LINK_TEXT,"Reports").click()
    
    time.sleep(5)
    driver.find_element(By.XPATH,"//button[@id='btnExportExcel']").click()
    time.sleep(5)
    # driver.find_element(By.LINK_TEXT ,"Reports").click()




if __name__ == "__main__":
    check=input("Type the required source: ('jira','replir) ")
    if check=="jira":
        jira()
    else:
        replir()




# driver.find_element(By.ID,"imp-ex-menu-container").click()
# time.sleep(3)
# driver.find_element(By.CLASS_NAME,"aui-list-section").click()                         
# x=driver.find_element(By.NAME,"allColumns")
# driver.implicitly_wait(3)
# drop=Select(x)
# drop.select_by_visible_text("All (max. 256 in xls)")
# win = driver.find_element(By.TAG_NAME,"html")
# win.send_keys(Keys.CONTROL + "-")
# win.send_keys(Keys.CONTROL + "-")
# win.send_keys(Keys.CONTROL + "-")

# time.sleep(2)
# driver.find_element(By.XPATH,"//button[contains(text(),'Export')]").click()
# time.sleep(80)
# driver.find_element(By.XPATH,"//a[contains(text(),'Download')]").click()

# driver.find_element(By.ID,"allExcelFields").click()

