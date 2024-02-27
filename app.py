from selenium import webdriver
from openpyxl import Workbook
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.common.exceptions import *
from time import sleep
import re

driver = webdriver.Chrome()
driver.maximize_window()

workbook = Workbook()
sheet = workbook.active

page_number = 0
url = "https://visitors.lineapelle-fair.it/#/event/lineapelle/catalogue/"
driver.get(url)
sleep(5)


def find_element(driver, whichBy, unique: str) -> WebElement:
    while True:
        try:
            element = driver.find_element(whichBy, unique)
            break
        except:
            pass
        sleep(3)
    return element


def find_elements(driver, whichBy, unique: str) -> list[WebElement]:
    while True:
        try:
            elements = driver.find_elements(whichBy, unique)
            break
        except:
            pass
        sleep(0.1)
    return elements


def products_items():
    element = ""
    while True:
        products_category = find_element(
            driver, By.XPATH, '/html/body/div[1]/div[1]/div/div/div/div[1]/div[2]/portlet/div/div[2]/div/div/div[1]/datatable-v2/div/div[2]/div[1]/div/div/div[2]/table/tbody').find_elements(By.XPATH, 'tr')
        item_amount = len(products_category)
        if item_amount:
            for i in range(item_amount):
                products_category = find_element(
                    driver, By.XPATH, '/html/body/div[1]/div[1]/div/div/div/div[1]/div[2]/portlet/div/div[2]/div/div/div[1]/datatable-v2/div/div[2]/div[1]/div/div/div[2]/table/tbody').find_elements(By.XPATH, 'tr')
                element1 = find_elements(products_category[i], By.TAG_NAME, 'td')[
                    0].find_element(By.TAG_NAME, 'label').text
                element2 = find_elements(products_category[i], By.TAG_NAME, 'td')[
                    1].find_element(By.TAG_NAME, 'label').text
                element += f"{element1} : {element2} \n"
            text = driver.find_element(
                By.XPATH, '/html/body/div[1]/div[1]/div/div/div/div[1]/div[2]/portlet/div/div[2]/div/div/div[1]/datatable-v2/div/div[2]/div[2]/div[1]/div[1]').text
            numbers = re.findall(r'\d+', text)
            if int(numbers[1]) == int(numbers[2]):
                break
            driver.find_element(
                By.XPATH, '/html/body/div[1]/div[1]/div/div/div/div[1]/div[2]/portlet/div/div[2]/div/div/div[1]/datatable-v2/div/div[2]/div[2]/div[2]/div/ul/li[5]').click()
            sleep(1)
        else:
            break
    return element


def get_info():
    try:
        driver.find_element(By.CLASS_NAME, 'm-card-profile__name')
        name = driver.find_element(By.CLASS_NAME, 'm-card-profile__name').text
    except:
        name = ""

    try:
        driver.find_element(
            By.CLASS_NAME, 'm-card-profile__email')
        address = driver.find_element(
            By.CLASS_NAME, 'm-card-profile__email').text
    except:
        address = ""
    try:
        driver.find_element(
            By.XPATH, '/html/body/div[1]/div[1]/div/div/div/div[1]/div[1]/div/div[2]/div/exhibitor-details/div[1]/div[2]/div[3]/div[2]/a')
        website = driver.find_element(
            By.XPATH, '/html/body/div[1]/div[1]/div/div/div/div[1]/div[1]/div/div[2]/div/exhibitor-details/div[1]/div[2]/div[3]/div[2]/a').text
    except:
        website = ""
    try:
        driver.find_element(
            By.XPATH, '/html/body/div[1]/div[1]/div/div/div/div[1]/div[1]/div/div[2]/div/exhibitor-details/div[1]/div[3]/div[2]/div[2]')
        description = driver.find_element(
            By.XPATH, '/html/body/div[1]/div[1]/div/div/div/div[1]/div[1]/div/div[2]/div/exhibitor-details/div[1]/div[3]/div[2]/div[2]').text
    except:
        description = ""
    try:
        driver.find_element(
            By.XPATH, '/html/body/div[1]/div[1]/div/div/div/div[1]/div[2]/portlet/div/div[2]/div/div/div[1]/datatable-v2/div/div[2]/div[1]/div/div/div[1]/div/table/thead/tr/th[1]')
        products = products_items()
    except:
        products = ""
    value = [name, address, website, description, products]
    sheet.append(value)
    workbook.save("output.xlsx")
    print(
        f"this is ===>{name}=======>{address}=======>{website}=====>{description}====>{products}")


while True:
    try:
        catalogues = find_element(
            driver, By.XPATH, '/html/body/div[1]/div[1]/div/div/div/portlet[2]/div/div[2]/div/div[2]/div/datatable-visitors/div/div[2]/div[1]/div/div[1]/div[2]/table/tbody').find_elements(By.TAG_NAME, 'tr')
        len_catalogues = len(catalogues)
        for i in range(len_catalogues):
            if page_number > 0:
                for j in range(page_number):
                    prenext = driver.find_element(
                        By.XPATH, "/ html/body/div[1]/div[1]/div/div/div/portlet[2]/div/div[2]/div/div[2]/div/datatable-visitors/div/div[2]/div[2]/div[2]/div/ul/li[9]")
                    driver.execute_script("arguments[0].click();", prenext)
                    sleep(1)
            catalogues = find_element(
                driver, By.XPATH, '/html/body/div[1]/div[1]/div/div/div/portlet[2]/div/div[2]/div/div[2]/div/datatable-visitors/div/div[2]/div[1]/div/div[1]/div[2]/table/tbody').find_elements(By.TAG_NAME, 'tr')
            element = find_elements(catalogues[i], By.TAG_NAME, 'td')[
                0].find_element(By.TAG_NAME, 'label')
            driver.execute_script("arguments[0].click();", element)
            sleep(5)

            # scrapping information
            get_info()
            driver.back()
            sleep(5)
        text = driver.find_element(
            By.XPATH, '/html/body/div[1]/div[1]/div/div/div/portlet[2]/div/div[2]/div/div[2]/div/datatable-visitors/div/div[2]/div[2]/div[1]/div[1]').text
        numbers = re.findall(r'\d+', text)
        if int(numbers[1]) == int(numbers[2]):

            break
        next = driver.find_element(
            By.XPATH, "/ html/body/div[1]/div[1]/div/div/div/portlet[2]/div/div[2]/div/div[2]/div/datatable-visitors/div/div[2]/div[2]/div[2]/div/ul/li[9]")
        driver.execute_script("arguments[0].click();", next)
        page_number += 1
        sleep(1)
    except Exception as e:
        pass
