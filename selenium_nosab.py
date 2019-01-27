from selenium import webdriver
import time
import xlwt

global record_num
global delay
global error_firma_name
record_num = 1
delay = 3  # seconds
driver = webdriver.Firefox()
book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("Sheet 1")
sheet1.write(0, 0, "Firma Adi")
sheet1.write(0, 1, "Firma Email")
sheet1.write(0, 2, "Firma Telefon")
sheet1.write(0, 3, "Firma Faks")
sheet1.write(0, 4, "Firma Web")
sheet1.write(0, 5, "Firma Adres")
sheet1.write(0, 6, "Firma Sektor")
sheet1.write(0, 7, "Vergi Numarasi")


def getInfo(row):
    global record_num
    try:
        # Firma bilgilerini alma islemi yapilacak.
        # /html/body/div[7]/div/div[2]/div[1]/div[1]
        firma_adi = driver.find_element_by_xpath("/html/body/div[7]/div/div[2]/div[1]").text
        firma_email = driver.find_element_by_xpath("/html/body/div[7]/div/div[2]/div[4]/div[4]").text
        firma_tel = driver.find_element_by_xpath("/html/body/div[7]/div/div[2]/div[4]/div[2]").text
        firma_faks = driver.find_element_by_xpath("/html/body/div[7]/div/div[2]/div[4]/div[3]").text
        firma_web = driver.find_element_by_xpath("/html/body/div[7]/div/div[2]/div[4]/div[5]/a").text
        firma_adres = driver.find_element_by_xpath("/html/body/div[7]/div/div[2]/div[4]/div[1]").text
        firma_sektor = driver.find_element_by_xpath("/html/body/div[7]/div/div[2]/div[2]/ul/li[2]").text
        firma_vergino = driver.find_element_by_xpath("/html/body/div[7]/div/div[2]/div[3]/ul[6]/li[2]").text

        # excel kayit alani
        sheet1.write(record_num, 0, firma_adi)
        sheet1.write(record_num, 1, firma_email)
        sheet1.write(record_num, 2, firma_tel)
        sheet1.write(record_num, 3, firma_faks)
        sheet1.write(record_num, 4, firma_web)
        sheet1.write(record_num, 5, firma_adres)
        sheet1.write(record_num, 6, firma_sektor)
        sheet1.write(record_num, 7, firma_vergino)
        book.save("nosab.xls")
        record_num += 1
        print("Kayit Edildi: " + firma_adi + " / " + firma_email + " / " + firma_tel[0:5] + " / " + firma_faks[0:5] + " / " + firma_web[0:5] + " / " + firma_adres[0:5] + " / " + firma_sektor[0:5] + " / " + firma_vergino[0:5])
        timeSleep(2, "getInfo")
    except Exception as e:
        print("Hata: getInfo: " + str(e))
        with open("errors.txt", "a+") as errorfile:
            errorfile.write("Nosab -- ")
            errorfile.write("Hata: getInfo: " + str(e))
            errorfile.write(" - Name : " + error_firma_name)
            errorfile.write("\n")
            errorfile.close()


def getCompanyPage(j):
    try:
        # /html/body/div[7]/div/div[2]/div[3]/ul/li[1]/a -> /html/body/div[7]/div/div[2]/div[3]/ul/li[2]/a
        error_firma_name = driver.find_element_by_xpath("/html/body/div[7]/div/div[2]/div[3]/ul/li[" + str(j) + "]/a").text
        driver.find_element_by_xpath("/html/body/div[7]/div/div[2]/div[3]/ul/li[" + str(j) + "]/a").click()
        print("Firma Sayfasina Gidildi: " + str(j))
        timeSleep(4, "getCompanyPage")
        # Firmanin bilgileri kayit edilecek
        getInfo(j)
        # Sayfada geri gidilip diger firmalarin bilgileri alinacak.
        getMainPage()
    except Exception as e:
        print("Hata: getCompanyPage: " + str(e))
        with open("errors.txt", "a+") as errorfile:
            errorfile.write("Nosab -- ")
            errorfile.write("Hata: getCompanyPage: " + str(e))
            errorfile.write(" - Name : " + error_firma_name)
            errorfile.write("\n")
            errorfile.close()


def getCompanies(numbers, page):
    try:
        for j in range(1, numbers + 1):
            getPage(page)  # Sayfalari sirayla cagirma islemi gerceklestiriliyor.
            getCompanyPage(j)
    except Exception as e:
        print("Hata: getCompanies: " + str(e))
        with open("errors.txt", "a+") as errorfile:
            errorfile.write("Nosab -- ")
            errorfile.write("Hata: getCompanies: " + str(e))
            errorfile.write(" - Name : " + error_firma_name)
            errorfile.write("\n")
            errorfile.close()


def getMainPage():
    try:
        driver.get("http://www.nosab.org.tr/firmalar/tr")
        print("Giris sayfasina gidildi.")
        timeSleep(4, "getMainPage")
    except Exception as e:
        print("Hata: getMainPage: " + str(e))
        with open("errors.txt", "a+") as errorfile:
            errorfile.write("Nosab -- ")
            errorfile.write("Hata: getMainPage: " + str(e))
            errorfile.write(" - Name : " + error_firma_name)
            errorfile.write("\n")
            errorfile.close()


def timeSleep(second, function):
    try:
        print("Sleep: " + str(second) + " Function: " + str(function))
        time.sleep(second)
    except Exception as e:
        print("Hata: timeSleep: " + str(e))
        with open("errors.txt", "a+") as errorfile:
            errorfile.write("Nosab -- ")
            errorfile.write("Hata: timeSleep: " + str(e))
            errorfile.write(" - Name : " + error_firma_name)
            errorfile.write("\n")
            errorfile.close()


def getNumbers():
    try:
        html_list = driver.find_element_by_xpath("/html/body/div[7]/div/div[2]/div[3]/ul")
        items = html_list.find_elements_by_tag_name("li")
        numbers = 1
        for item in items:
            numbers += 1
        return numbers
    except Exception as e:
        print("Hata: getNumbers: " + str(e))
        with open("errors.txt", "a+") as errorfile:
            errorfile.write("Nosab -- ")
            errorfile.write("Hata: getNumbers: " + str(e))
            errorfile.write(" - Name : " + error_firma_name)
            errorfile.write("\n")
            errorfile.close()


def getPage(page):
    try:
        # /html/body/div[7]/div/div[2]/ul/li[1]/a
        driver.find_elements_by_xpath("/html/body/div[7]/div/div[2]/ul/li[" + str(page) + "]/a")[0].click()
        print("Sayfa: " + str(page))
        timeSleep(4, "getPage")
    except Exception as e:
        print("Hata: getPage: " + str(e))
        with open("errors.txt", "a+") as errorfile:
            errorfile.write("Nosab -- ")
            errorfile.write("Hata: getPage: " + str(e))
            errorfile.write(" - Name : " + error_firma_name)
            errorfile.write("\n")
            errorfile.close()


def main():
    try:
        getMainPage()
        for i in range(1, 30):  # A-Z olan sayfalar her firma listesi bittiginde diger sayfaya gecilecek.
            numbers = getNumbers()  # Girilen sayfadaki firma sayfasi donuyor
            print("Total Company This Page: " + str(getNumbers()))
            getCompanies(numbers, i)
        driver.close()
    except Exception as e:
        print("Hata: main: " + str(e))
        with open("errors.txt", "a+") as errorfile:
            errorfile.write("Nosab -- ")
            errorfile.write("Hata: main: " + str(e))
            errorfile.write(" - Name : " + error_firma_name)
            errorfile.write("\n")
            errorfile.close()


main()
