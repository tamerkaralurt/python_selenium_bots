from selenium import webdriver
import time
import xlwt
import re

global record_num
record_num = 1

driver = webdriver.Firefox()
driver.get("http://kayseriosb.org/tr/5/Firmalar.html")
time.sleep(1)
book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("Sheet 1")
sheet1.write(0, 0, "Firma Adi")
sheet1.write(0, 1, "Firma Email")
sheet1.write(0, 2, "Firma Telefon")
sheet1.write(0, 3, "Firma Faks")
sheet1.write(0, 4, "Firma Web")
sheet1.write(0, 5, "Firma Adres")
sheet1.write(0, 6, "Firma Sektor")


def getInfo(row):
    global record_num
    try:
        # Firma bilgilerini alma islemi yapilacak.
        # //*[@id="main"]/div/div/article/div[2]/div/div[2]/div/h2
        # print("row: " + str(row))
        firma_adi = driver.find_elements_by_xpath("//*[@id='main']/div/div/article/div[2]/div/div[2]/div/h2")[0].text
        firma_email = driver.find_elements_by_xpath("//*[@id='main']/div/div/article/div[2]/div/div[2]/ul/li[6]")[0].text
        firma_tel = driver.find_elements_by_xpath("//*[@id='main']/div/div/article/div[2]/div/div[2]/ul/li[4]")[0].text
        firma_faks = driver.find_elements_by_xpath("//*[@id='main']/div/div/article/div[2]/div/div[2]/ul/li[5]")[0].text
        firma_web = driver.find_elements_by_xpath("//*[@id='main']/div/div/article/div[2]/div/div[2]/ul/li[7]")[0].text
        firma_adres = driver.find_elements_by_xpath("//*[@id='main']/div/div/article/div[2]/div/div[2]/ul/li[2]")[0].text
        firma_sektor = driver.find_elements_by_xpath("//*[@id='main']/div/div/article/div[2]/div/div[2]/ul/li[1]")[0].text

        # parcalama islemleri
        firma_email = firma_email.split(": ")
        firma_tel = firma_tel.split(": ")
        firma_faks = firma_faks.split(": ")
        firma_web = firma_web.split(": ")
        firma_adres = firma_adres.split(": ")
        firma_sektor = firma_sektor.split(": ")

        # parca kontrol alani
        if len(firma_email) == 1:
            firma_email = ""
        else:
            firma_email = str(firma_email[1]).lower()

        if len(firma_tel) == 1:
            firma_tel = ""
        else:
            firma_tel = str(firma_tel[1])

        if len(firma_faks) == 1:
            firma_faks = ""
        else:
            firma_faks = str(firma_faks[1])

        if len(firma_web) == 1:
            firma_web = ""
        else:
            firma_web = str(firma_web[1])

        if len(firma_adres) == 1:
            firma_adres = ""
        else:
            firma_adres = str(firma_adres[1])

        if len(firma_sektor) == 1:
            firma_sektor = ""
        else:
            firma_sektor = str(firma_sektor[1]).lower()

        # excel kayit alani
        sheet1.write(record_num, 0, firma_adi)
        sheet1.write(record_num, 1, firma_email)
        sheet1.write(record_num, 2, firma_tel)
        sheet1.write(record_num, 3, firma_faks)
        sheet1.write(record_num, 4, firma_web)
        sheet1.write(record_num, 5, firma_adres)
        sheet1.write(record_num, 6, firma_sektor)
        book.save("kayseri.xls")
        record_num += 1
        timeSleep(2)
        getBack()
    except Exception as e:
        print("getInfo: " + str(e))
        row += 1
        getBack()


def getCategory(category):
    try:
        print("Kategori: " + str(category))
        driver.find_elements_by_xpath("//*[@id='main']/div/div/article/div[2]/div[2]/div[4]/ul/li[" + str(category) + "]/a")[0].click()
        timeSleep(2)
    except Exception as e:
        print("getCategory: " + str(e))


def getBack():
    try:
        driver.forward()
        driver.back()
    except Exception as e:
        print("getBack: " + str(e))


def timeSleep(second):
    time.sleep(second)


def getNumbers():
    try:
        info = driver.find_element_by_id("table_info")
        info = info.get_attribute('innerHTML')
        numbers = re.findall(r'\d+', info)
        return numbers
    except Exception as e:
        print("getNumbers: " + str(e))

def getMainPage():
    driver.get("http://kayseriosb.org/tr/5/Firmalar.html")


def getPage(page):
    try:
        driver.find_elements_by_xpath("//*[@id='table_paginate']/ul/li[" + str(page) + "]/a")[0].click()
        timeSleep(1)
    except Exception as e:
        print("getPage: " + str(e))


def recordCompanies(page, record):
    numbers = getNumbers()
    total_records = numbers[0]
    if int(record) == int(total_records):
        return
    row = 1
    for j in range(int(numbers[1]), int(numbers[2]) + 1):
        try:
            print("Firma: " + str(j) + " (" + numbers[0] + " - " + numbers[1] + " / " + numbers[2] + ")")
            driver.find_elements_by_xpath("//*[@id='tableData']/tr[" + str(row) + "]/td[1]/a")[0].click()
            timeSleep(2)
            getInfo(j)
            row += 1
            getPage(page)
            if (int(numbers[2]) == j) and (int(j) <= int(total_records)):
                row = 1
                page += 1
                getPage(page)
                recordCompanies(page, j)
        except Exception as e:
            print("recordCompanies: " + str(e))


def main():
    try:
        for i in range(1, 13):
            getCategory(i)
            page = 2
            record = 1
            recordCompanies(page, record)
            getMainPage()
            timeSleep(2)

        driver.close()
    except Exception as e:
        print("main: " + str(e))


main()
