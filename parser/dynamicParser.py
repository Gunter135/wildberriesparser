from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import openpyxl

#url = "https://www.wildberries.ru/brands/solgar"
url = str(input("Введите URL "))
pages = int(input("Введите количество страниц для парсинга "))
driver = webdriver.Chrome(executable_path="chromedriver.exe")

try:
    driver.get(url)
    time.sleep(2)
    book = openpyxl.Workbook()
    for sheet in book.worksheets:
        sheet.insert_rows(0)
        sheet["A1"].value = "Стоимость товара"
        sheet["B1"].value = "Производитель"
        sheet["C1"].value = "Наименование товара"
        sheet.column_dimensions["A"].width = 20
        sheet.column_dimensions["B"].width = 20
        sheet.column_dimensions["C"].width = 100
    for i in range(pages):
        for i in range(70):
            driver.execute_script("window.scrollTo(0, 20000)")
        data = driver.find_elements(By.CLASS_NAME, "product-card.product-card--hoverable.j-card-item")

        for d in data:
            arr = d.text.split("\n")
            if 'NEW' in arr:
                arr.remove('NEW')
            if 'яГОДНЫЕ СКИДКИ' in arr:
                arr.remove('яГОДНЫЕ СКИДКИ')
            if '%' in arr[0]:
                arr.remove(arr[0])
            st = ""
            if arr[0].find('₽') == (len(arr[0]) - 1):
                for i in range(len(arr[0])):
                    if arr[0][i] != ' ' and arr[0][i] != '₽':
                        st += arr[0][i]
                arr[0] = st
            else:
                for i in range(arr[0].find('₽'), len(arr[0])):
                    if arr[0][i] != ' ' and arr[0][i] != '₽':
                        st += arr[0][i]
                arr[0] = st
            arr = arr[:3:]
            for sheet in book.worksheets:
                sheet.append(arr)
        time.sleep(1)
        driver.find_element(By.CLASS_NAME, "pagination-next.pagination__next.j-next-page").click()
        time.sleep(1)
    book.save("Prices.xlsx")
    time.sleep(1)

except Exception as ex:
    print(ex)
finally:
    driver.close()
    driver.quit()
