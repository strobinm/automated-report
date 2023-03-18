from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import datetime as dt
import xlwings as xw


options = webdriver.FirefoxOptions()
options.headless = True
PSHandover = xw.Book('ICQA PS HANDOVER testy.xlsm')
handoverSheet = PSHandover.sheets['PS Handover']
#  badgeBarcodesSheet = PSHandover.sheets["Sheet1"]
date = handoverSheet['B1'].options(dates=dt.date).value
date = str(date)
# create a new Firefox instance
driver = webdriver.Firefox(options=options)
resolvedAndons = "http://example.com/LCJ2?category=Bin+Item+Defects&type=All+types&status=Resolved&startDate=" + date
appLuncher = "https://example.com"
# navigate to appLuncher Login
driver.get(appLuncher)
# get input box xPath
xPath = "/html/body/div[2]/div/div[3]/form/input"
#   locate the search box element
search_box = driver.find_element(by=By.XPATH, value=xPath)
badgeBarcodeId = '1234'
# type "badgeBarcodeId" in the search box
search_box.send_keys(badgeBarcodeId)
# press the Enter key
search_box.send_keys(Keys.RETURN)
# wait until page will load
time.sleep(1)
# navigate to FC Andons
driver.get(resolvedAndons)
# xPath to CSV file download button
dowloadCSVButton = "/html/body/div/div/div/awsui-app-layout/div/main/div/div[2]/div/span/div/awsui-table/div/div[2]/div[1]/div[1]/span/div/div[2]/awsui-button[1]/button/span"
# wait until page is loaded
time.sleep(2)
# click download CSV button
search_box = driver.find_element(by=By.XPATH, value=dowloadCSVButton).click()
# close Firefox window
driver.close()
print("done")
