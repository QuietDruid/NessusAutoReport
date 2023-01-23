import time
import pandas as pd
import os
import glob
from pathlib import Path
import selenium.webdriver.support.ui as ui
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as ec
import pdfmaker
import folderhandler

# following opens Nessus and runs through the steps of traversing multiple forms to download the csv
# waits for certain form items to load because some are not immediately detected and at the end
# waits for csv to download as some computers may run slow


def handlecsvdownload():
    # initialization of logging in
    options = webdriver.EdgeOptions()
    options.add_argument('--ignore-ssl-errors=yes')
    options.add_argument('--ignore-certificate-errors=yes')

    driver = webdriver.Edge(options=options)
    driver.get('https://localhost:8834/')

    # Form Handling
    wait = ui.WebDriverWait(driver, 15)
    element = wait.until(ec.visibility_of_element_located((By.XPATH, "/html/body/div/form/div[1]/input")))
    element.send_keys('username') #insert username used
    element = wait.until(ec.visibility_of_element_located((By.XPATH, "/html/body/div/form/div[2]/input")))
    element.send_keys('password') #insert nessus password
    element.send_keys('\ue007')  # enter

    # dealing with scans in list
    # CSS tr.scan:nth-child(1)
    element = wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'tbody:nth-child(2)')))
    num = element.find_elements(By.TAG_NAME, 'tr')

    # prints scan list
    for table in num:
        table = table.find_element(By.CSS_SELECTOR, "td:nth-child(3)")
        print("\nCurrent Scan Being Created: " + table.text)

    # first scan in list
    firstScan = element.find_element(By.CSS_SELECTOR, 'tr.scan:nth-child(1)')
    firstScan.click()

    element = wait.until(ec.visibility_of_element_located((By.ID, "generate-scan-report")))
    element.click()

    element = wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, "div.radio:nth-child(5)")))
    element.click()

    element = wait.until(ec.visibility_of_element_located((By.ID, "report-save")))
    element.click()
    time.sleep(10)
    driver.quit()

# weird division of work with various functions but further refactoring to come down the line if to keep maintainable
# checks if folders exists, then downloads csv, then runs through formatting and pdf creation before cleaning up
# downloads folder


if __name__ == '__main__':
    folderhandler.doesfolderexist()

    handlecsvdownload()

    pdfmaker.callpdfmaker()

    folderhandler.movetofolder()
