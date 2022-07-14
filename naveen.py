import time
import unittest
import HtmlTestRunner
from selenium import webdriver
from selenium.webdriver.common.by import By
import EXCELUTIL
import openpyxl

#class DemoImplicitWait():
class Testsample(unittest.TestCase):
    def test_amazon(self):
        self.driver = webdriver.Chrome(executable_path="C:/chromedrive/chromedriver.exe")

        self.driver.implicitly_wait(20)



        self.driver.get("https://www.amazon.in")
        titleOfWebPage= self. driver.title
        self.assertNotEqual("Amazon",titleOfWebPage,"PASSED")
          #seconds steps

        #to maximize the window
        self.driver.maximize_window()

        # capture the title of the page
        amazon = self.driver.title
        path = ("C:\\Users\\Admin\\OneDrive\\Desktop\\Naveen.xlsx")

        # access the RowCount method
        rows = EXCELUTIL.getRowCount(path, "Sheet1")
        # perform or read the value from excel file and pass to application

        for r in range(2, rows + 1):
            Actionkeywords = EXCELUTIL.ReadData(path, "Sheet1", r, 1)
            status = EXCELUTIL.ReadData(path, "Sheet1", r, 2)

        #time.sleep(4)
        self.driver.find_element(by=By.XPATH, value=" //input[@id='twotabsearchtextbox']").send_keys("loptops")

        # screenshot
        self.driver.save_screenshot("C:\screenshot\search loptop .jpeg")
        time.sleep(4)
        self.driver.find_element(by=By.XPATH, value=" //input[@id='nav-search-submit-button']").click()






        self.driver.find_element(by=By.XPATH, value='//*[@id="a-autoid-0-announce"]').click()
       # time.sleep(4)
        self.driver.find_element(by=By.XPATH, value='//*[@id="s-result-sort-select_2"]').click()

        # screenshot
        self.driver.save_screenshot("C:\screenshot\sort by feature.jpeg")

        time.sleep(4)


        self.driver.execute_script("alert('search successfully');")
        time.sleep(4)


        self.driver.switch_to.alert.dismiss()

        workbook = openpyxl.load_workbook("C:\\Users\\Admin\\OneDrive\\Desktop\\Naveen.xlsx")

        # load of the sheet
        sheet = workbook.active

        for r in range(1, 6):
            for c in range(1, 2):
                sheet.cell(row=r, column=c).value = "pass "

        workbook.save("C:\\Users\Admin\\OneDrive\\Desktop\\Naveen.xlsx")

        print("end of file writting")

        # time.sleep(4)

        self.driver.close()

if __name__=="__main__":
    unittest.main(testRunner=HtmlTestRunner.HTMLTestRunner(output="C:\\Users\\Admin\\CGunittest\\RPT"))