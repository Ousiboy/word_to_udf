import os
import subprocess
import time
import pyautogui
import sys
import keyboard
import pyperclip
import tkinter as tk
from tkinter import filedialog, Tk, messagebox
from pathlib import Path
from datetime import datetime

desktop_path_str = ""  # desktop_path_str'ı global bir değişken olarak tanımlayın

def kopyala():
    keyboard.press("ctrl")
    keyboard.press("a")
    keyboard.release("a")
    keyboard.release("ctrl")
    keyboard.press("ctrl")
    keyboard.press("c")
    keyboard.release("c")
    keyboard.release("ctrl")

def fotokonum(image_name):
    image_path = f"{image_name}.png"

    # Görüntüyü bul
    image_location = pyautogui.locateOnScreen(image_path, confidence=0.9)

    if image_location is not None:
        # Görüntünün merkezini al
        image_center = pyautogui.center(image_location)

        # Fareyi görüntünün merkezine taşı
        pyautogui.moveTo(image_center, duration=0.40)

    else:
        print(f"{image_path} görüntüsü bulunamadı.")
        sys.exit()

def open_document():
    root = tk.Tk()
    root.withdraw()

    # Kullanıcıya dosya seçme penceresini gösterin
    file_path = filedialog.askopenfilename()

    if not file_path:
        print("Dosya seçilmedi.")
        return

    # Dosyayı aç
    try:
        subprocess.Popen(["start", "", file_path], shell=True)
    except subprocess.CalledProcessError as e:
        print("Dosya açılamadı:", e)
        return

    print("Dosya başarıyla açıldı:", file_path)

    # Dosya adını ve dizinini al
    file_name = os.path.basename(file_path)
    directory = os.path.dirname(file_path)

    # Klasörü Masaüstü altında oluştur
    desktop_path = None

    # 1. Adım: OneDrive/Masaüstü kontrolü
    onedrive_path = Path.home() / "OneDrive" / "Masaüstü"
    if onedrive_path.exists():
        desktop_path = onedrive_path

    # 2. Adım: Ana dizinde "masaüstü" kontrolü
    home_path = Path.home()
    masaustu_path = home_path / "Masaüstü"
    if desktop_path is None and masaustu_path.exists():
        desktop_path = masaustu_path

    # 3. Adım: Ana dizinde "desktop" kontrolü
    desktop_path = home_path / "Desktop" if desktop_path is None else desktop_path

    # word_to_udf klasörünü oluştur
    word_to_udf_path = desktop_path / "word_to_udf"

    try:
        os.makedirs(word_to_udf_path, exist_ok=True)
        print("word_to_udf klasörü başarıyla oluşturuldu.")
    except OSError as e:
        print("word_to_udf klasörü oluşturulamadı:", e)
        sys.exit()

    desktop_path_str = str(desktop_path / "word_to_udf")  # word_to_udf klasörünün tam yolunu oluşturuyoruz

    time.sleep(3)  # Dosyanın açılmasını bekleyin

    # Microsoft Office Etkinleştirme Sihirbazı penceresini kapatın
    time.sleep(5)  # Pencerenin ortaya çıkmasını bekleyin
    sihirbaz_penceresi = pyautogui.getWindowsWithTitle("Microsoft Office Etkinleştirme Sihirbazı")

    if sihirbaz_penceresi:
        sihirbaz_penceresi[0].close()
        print("Microsoft Office Etkinleştirme Sihirbazı penceresi kapatıldı.")
    else:
        pass

    # Düzenlemeyi etkinleştir tıkla
    time.sleep(1)  # Kendine gelmesini bekle
    dznlmeetknlstr = pyautogui.locateOnScreen('dznlmtknlstr.png')

    if dznlmeetknlstr is not None:
        dznlmtknlstrorta = pyautogui.center(dznlmeetknlstr)
        pyautogui.moveTo(dznlmtknlstrorta, duration=0.2)
        pyautogui.click()
        time.sleep(1)
    else:
        pass

    # Dosyayı tam ekran olarak aç
    time.sleep(1)  # Tam ekran geçişini bekleyin
    word_penceresi = pyautogui.getWindowsWithTitle(os.path.basename(file_path))[0]
    if word_penceresi:
        word_penceresi.maximize()
        print("Word dosyası tam ekran olarak açıldı.")

    # Farklı kaydet sayfasını aç
    pyautogui.press("f12")
    print("Farklı kaydet sayfası açıldı.")

    time.sleep(2)  # Dosyanın açılmasını bekleyin

    # Kayıt türü seç
    fotokonum("kayit")
    pyautogui.move(110, 0, 0.5)
    pyautogui.click()
    time.sleep(0.5)

    # .rtf seç
    fotokonum("zngn")
    pyautogui.click()

    # kaydetme yeri olarak Masaüstü seç
    fotokonum("kydtmyr")
    pyautogui.move(-150, 0, 0.3)
    time.sleep(0.1)
    pyautogui.click()
    keyboard.write(desktop_path_str)  # dize olarak yazdırıyoruz
    pyautogui.press('enter')
    pyautogui.click()

    # .isim satırına tıkla
    fotokonum("dosyaadi")
    pyautogui.move(700, 0, 1)
    pyautogui.click()
    pyautogui.click()
    pyautogui.click()

    # isim satırını sil
    pyautogui.hotkey("backspace", "backspace", "backspace", "backspace", interval=0.10)
    kopyala()
    global dosya_adi
    time.sleep(0.5)
    dosya_adi = pyperclip.paste()
    pyautogui.move(-750, 0, 0.4)
    pyautogui.move(0, 120, 0.2)
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(0.9)

    # dosya zaten var penceresi denetle
    count = 0
    while count < 5:
        tamambutton = pyautogui.locateOnScreen('tmm.png', confidence=0.8)
        if tamambutton is not None:
            time.sleep(0.2)
            pyautogui.press('down')
            time.sleep(0.2)
            pyautogui.press('enter')
            time.sleep(1.5)
            fotokonum("dosyaadi")
            pyautogui.move(700, 0, 0.5)
            pyautogui.click()
            time.sleep(0.2)
            pyautogui.click()
            time.sleep(0.2)
            pyautogui.click()
            time.sleep(0.2)
            pyautogui.write("_" + str(count+1))
            time.sleep(0.5)
            kopyala()
            time.sleep(0.8)
            dosya_adi = pyperclip.paste()
            time.sleep(0.3)
            pyautogui.press('enter')
            print("Aynı isim ve uzantıda bulunan dosya mevcut.")
            print(f"Dosya ismi '{dosya_adi}' olarak kaydedildi.")
            count = count+1
            time.sleep(1.5)
        else:
            print("Aynı isimli dosya yok.")
            break
    
    # kapat tıkla
    pyautogui.hotkey("alt", "f4")

    time.sleep(1)
    return desktop_path_str  # desktop_path_str değerini döndürün
    time.sleep(2)
# open_document() fonksiyonunu çağırın ve sonucunu desktop_path_str'ye atayın
desktop_path_str = open_document()

# İkinci programı burada çalıştırın
import os
from tkinter import messagebox, Tk, filedialog

def run_program():
    try:
        os.startfile("C:\\Uyap\\Uyap Kelime Islemci\\Uyap Doküman Editörü x86.exe")
    except FileNotFoundError:
        try:
            os.startfile("C:\\Uyap\\Uyap Kelime Islemci\\Uyap Doküman Editörü.exe")
        except FileNotFoundError:
            try:
                os.startfile("D:\\Uyap\\Uyap Kelime Islemci\\Uyap Doküman Editörü x86.exe")
            except FileNotFoundError:
                try:
                    os.startfile("D:\\Uyap\\Uyap Kelime Islemci\\Uyap Doküman Editörü.exe")
                except FileNotFoundError:
                    root = Tk()
                    root.withdraw()
                    messagebox.showinfo("Dosya Bulunamadı", "Uyap Doküman Editörü programını bulamadım. Lütfen programın bulunduğu konumu seçin.")
                    file_path = filedialog.askopenfilename(title="Uyap Doküman Editörü Programını Seçin")
                    if file_path:
                        os.startfile(file_path)

run_program()

time.sleep(5)  # Uyap Doküman Editörü'nün yüklenmesini bekleyin (süreyi isteğinize göre ayarlayın)

# Tam ekran moduna geçiş yapmak için Uyap Doküman Editörü'nün penceresini bulun
while True:
    uyap_penceresi = pyautogui.getWindowsWithTitle("Doküman Editörü v5.4.6 - isimsiz.UDF")
    if uyap_penceresi:
        uyap_penceresi[0].maximize()
        print("Uyap Doküman Editörü üzerinde işlem yapmaya başlandı.")
        break
    time.sleep(3)

# Open sayfasını aç

pyautogui.hotkey('ctrl', 'o', confidence=0.8)
time.sleep(2)  # "Open" penceresinin açılmasını bekle

# dosya türü seç
fotokonum("oftypebutton")
pyautogui.move(100, 0, 0.4)
pyautogui.click()
time.sleep(0.8)

# rtf türü seç
fotokonum("rtfdosyasi")
time.sleep(0.2)
pyautogui.click()

# kaydetme yeri seç
fotokonum("lookin")
pyautogui.move(330, 0, 0.5)
time.sleep(0.2)
pyautogui.click(clicks=10, interval=0.15)

# ilk dosyayı seç
fotokonum("openbutton")
pyautogui.move(-30, 80, 0.4)
pyautogui.click()
time.sleep(0.3)

# word_to_udf dosyasını seç

for _ in range(10):
    pyautogui.press('w')
    time.sleep(0.2)  # 0.2 saniye bekle
    kopyala()
    time.sleep(0.2)
    file_name = pyperclip.paste()  # Dosya adını al

    if file_name.strip() == desktop_path_str:
        pyautogui.press('enter')
        print("Seçilen dosya 'word_to_udf' klasörüne kaydedildi.")
        break  # Döngüden çık
    else:
        print(f"Seçilen dosya '{file_name}' değil.")
else:
    print("Dosya adı algılanamadı.")
    sys.exit()

# ilk dosyayı seç
fotokonum("openbutton")
pyautogui.move(-30, 80, 0.4)
pyautogui.click()
time.sleep(0.3)

# kaydedillen rtf dosyasını seç
aranandosya = desktop_path_str + "\\" + dosya_adi + ".rtf"
for _ in range(10):
    time.sleep(0.5)  # 0.2 saniye bekle
    kopyala()
    time.sleep(0.5)
    rtf_name = pyperclip.paste()  # Dosya adını al

    if rtf_name == aranandosya:
        pyautogui.press('enter')
        print("Dosya seçildi ve kaydedildi.")
        break  # Döngüden çık
    else:
        print(f"Seçilen dosya '{rtf_name}' değil.")
        print(f"Seçilen dosya '{aranandosya}'.")
        pyautogui.press('down')
else:
    print("Dosya adı algılanamadı.")
    sys.exit()
    
time.sleep(1)
# hepsini seç
pyautogui.hotkey('ctrl', 'a')
time.sleep(0.3)

# ters tıkla
pyautogui.click(button='right')
time.sleep(0.5)

# Yazı özellikleri menüsünü seç
fotokonum("yaziozellikleri")
pyautogui.click()
time.sleep(0.5)

# Yazı türü kısmına tıkla
fotokonum("yazitipiadi")
pyautogui.move(80, 18, 0.5)
pyautogui.click()

# Times New Roman seç
fotokonum("tmsnwrmn")
pyautogui.click(clicks=2)

# Yazı boyutu kısmına tıkla
fotokonum("yaziboyutuadi")
pyautogui.move(30, 17, 0.5)
pyautogui.click()

# Yazı boyutu seç
fotokonum("oniki")
pyautogui.click(clicks=2)

# Tamam tespit et ve tıkla
tamambutonu = pyautogui.locateOnScreen('yazitamambutton.png')
yazitamambutonorta = pyautogui.center(tamambutonu)
pyautogui.moveTo(yazitamambutonorta)
pyautogui.click()
time.sleep(0.5)

# Sayfanın üstüne çık
pyautogui.click()
pyautogui.press('pageup', presses=30)
time.sleep(0.3)

# Farklı kaydet kısayol aç ve isim kısmına git.
pyautogui.hotkey('shift', 'ctrl', 's')
time.sleep(0.5)
fotokonum("udfbutton")
pyautogui.move(30, -28, 0.3)
pyautogui.click()
time.sleep(0.2)
pyautogui.press("backspace", presses=4)

# Save kısmını algıla ve tıkla.
fotokonum("savebutton")
pyautogui.click()
time.sleep(0.5)
print("Çevirme işlemi başarıyla tamamlandı.")
