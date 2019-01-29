## Python ile selenium bot
- Python ve selenium ile sitelerden alınan bilgilerin excel dosyasına yazılması

## Kurulum
- Download - [python-3.7.2.exe](https://www.python.org/ftp/python/3.7.2/python-3.7.2.exe)
- Download - [Firefox](https://www.mozilla.org/tr/firefox/download/thanks/)
- Download - [Gecko Driver](https://github.com/mozilla/geckodriver/releases)

- **İndirdikten sonra kurulumlari gerçekleştirin**

- Gecko kurulmaz ancak .exe dosyasını bilgisayarınızın path alanına eklemeniz gerekmektedir.
    - Windows 10 ve Windows 8
        1. Ara'da Sistem (Denetim Masası) öğesini arayın ve seçin
        2. Gelişmiş sistem ayarları bağlantısına tıklayın.
        3. Ortam Değişkenleri'ne tıklayın. Sistem Değişkenleri bölümünde PATH ortam değişkenini bulup seçin. Düzenle'ye tıklayın. PATH ortam değişkeni yoksa Yeni'ye tıklayın.
        4. Sistem Değişkenini Düzenle (veya Yeni Sistem Değişkeni) penceresinde PATH ortam değişkeninin değerini belirtin. Tamam'a tıklayın. Tamam'a tıklayarak kalan tüm pencereleri kapatın.
            - Örnek: C:\Users\uzman\Desktop\selenium_python

- **Konsola Kodu Girin - Selenium Kurulumu**
    - pip install selenium