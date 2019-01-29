from selenium import webdriver
import time
import xlwt

global record_num
record_num = 1
driver = webdriver.Firefox()
book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("Sheet 1")
sheet1.write(0, 0, "Firma Adi")
sheet1.write(0, 1, "Firma Email")
sheet1.write(0, 2, "Firma Telefon")
sheet1.write(0, 3, "Firma Faks")
sheet1.write(0, 4, "Firma Web")


def getInfo(row):
    global record_num
    try:
        # Firma bilgilerini alma islemi yapilacak.
        # /html/body/div/div[1]/form/div[7]/div/div/div/div/div[2]/div[2]/div/h2/a
        firma_adi = driver.find_element_by_xpath("/html/body/div/div[1]/form/div[7]/div/div/div/div/div["+str(row)+"]/div[2]/div/h2/a").text
        firma_email = driver.find_element_by_xpath("/html/body/div/div[1]/form/div[7]/div/div/div/div/div["+str(row)+"]/div[2]/div/div[1]/ul/li[4]").text
        firma_tel = driver.find_element_by_xpath("/html/body/div/div[1]/form/div[7]/div/div/div/div/div["+str(row)+"]/div[2]/div/div[1]/ul/li[2]").text
        firma_faks = driver.find_element_by_xpath("/html/body/div/div[1]/form/div[7]/div/div/div/div/div["+str(row)+"]/div[2]/div/div[1]/ul/li[3]").text
        firma_web = driver.find_element_by_xpath("/html/body/div/div[1]/form/div[7]/div/div/div/div/div["+str(row)+"]/div[2]/div/div[1]/ul/li[1]").text

        # excel kayit alani
        sheet1.write(record_num, 0, firma_adi)
        sheet1.write(record_num, 1, firma_email)
        sheet1.write(record_num, 2, firma_tel)
        sheet1.write(record_num, 3, firma_faks)
        sheet1.write(record_num, 4, firma_web)
        book.save("Corum.xls")
        record_num += 1
        print("Kayit Edildi: " + firma_adi + " / " + firma_email + " / " + firma_tel[0:5] + " / " + firma_faks[0:5] + " / " + firma_web[0:5])
        timeSleep(2, "getInfo")
    except Exception as e:
        print("Hata: getInfo: " + str(e))
        with open("errors.txt", "a+") as errorfile:
            errorfile.write("Corum -- ")
            errorfile.write("Hata: getInfo: " + str(e))
            errorfile.write("\n")
            errorfile.close()


def getCompanies():
    # /html/body/div/div[1]/form/div[7]/div/div/div/div/div[1]
    # /html/body/div/div[1]/form/div[7]/div/div/div/div/div[2]
    items = driver.find_elements_by_class_name('blog-item')
    for item in items:
        getInfo(items.index(item)+1)

def getMainPage():
    try:
        driver.get("https://www.corumosb.org.tr/FirmaListele.aspx")
        print("Giris sayfasina gidildi.")
        print("##########################")
        timeSleep(4, "getMainPage")
    except Exception as e:
        print("Hata: getMainPage: " + str(e))
        with open("errors.txt", "a+") as errorfile:
            errorfile.write("Corum -- ")
            errorfile.write("Hata: getMainPage: " + str(e))
            errorfile.write("\n")
            errorfile.close()


def timeSleep(second, function):
    try:
        # print("Sleep: " + str(second) + " Function: " + str(function))
        time.sleep(second)
    except Exception as e:
        print("Hata: timeSleep: " + str(e))
        with open("errors.txt", "a+") as errorfile:
            errorfile.write("Corum -- ")
            errorfile.write("Hata: timeSleep: " + str(e))
            errorfile.write("\n")
            errorfile.close()


def main():
    try:
        getMainPage()
        getCompanies()
        driver.close()
        exit()
    except Exception as e:
        print("Hata: main: " + str(e))
        with open("errors.txt", "a+") as errorfile:
            errorfile.write("Corum -- ")
            errorfile.write("Hata: main: " + str(e))
            errorfile.write("\n")
            errorfile.close()


main()