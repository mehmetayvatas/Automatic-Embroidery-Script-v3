from os import times
import subprocess
import pyautogui
import pyodbc
import time
import pyperclip
import os
import shutil

pyautogui.FAILSAFE = False

def Programi_Baslat():
    program_yolu = r"C:\Program Files (x86)\G7Solutions\FlorianiTotalControl\FlorianiPro.exe"
    subprocess.Popen(program_yolu, shell=True)
    time.sleep(10)
    pyautogui.click(x=641, y=743) # Programda Acilis Koordinatlari
    pyautogui.click(x=23, y=92)
    pyautogui.click(x=659, y=477)

def noktalari_kaldir(text):
    return text.replace('.', ' ').replace('/', ' ')

def SQL_Baglantisi():
    conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=10.1.10.2;DATABASE=URUNYONETIMI;UID=sa;PWD=Sa1234')
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM MONOGRAM_DETAY')

    onceki_sipno_font_tur = None
    onceki_font = None  

    for row in cursor.fetchall():
        if not row.YAPILDI:
            if row.FONT not in ["Angelic Curls", "Athletic", "Arial", "Bamboo", "Bauhaus", "Benquiat", "Bordeaux", "Brush", "Diana-Vs", "Freehand", "Giddyup", "Impress", "Jester Pro", "Kids", "Manila", "Old_Engl", "Pepita", "Ribbon Love", "Swiss", "Tango", "Times", "Trad_Scr", "Universi"]:
                cursor.execute("UPDATE MONOGRAM_DETAY SET YAPILDI = ? WHERE ID = ?", (True, row.ID))
                conn.commit()
                continue

            pyperclip.copy(row.YAZI)
            pyautogui.click(x=1462, y=177)  # Metin Alani Koordinatlari
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.hotkey('ctrl', 'v')
            pyautogui.click(x=1676, y=383)

            font_harfleri = {
                "Angelic Curls": ("a", 1),
                "Arial": ("a", 4),
                "Athletic": ("a", 10),
                "Bamboo": ("b", 0),
                "Bauhaus": ("b", 1),
                "Benquiat": ("b", 3),
                "Bordeaux": ("b", 10),
                "Brush": ("b", 14),
                "Diana-Vs": ("d", 5),
                "Freehand": ("f", 8),
                "Giddyup": ("g", 1),
                "Impress": ("i", 1),
                "Jester Pro": ("j", 1),
                "Kids": ("k", 1),
                "Manila": ("m", 0),
                "Old_Engl": ("o", 2),
                "Pepita": ("p", 2),
                "Ribbon Love": ("r", 3),
                "Swiss": ("s", 12),
                "Tango": ("t", 1),
                "Times": ("t", 3),
                "Trad_Scr": ("t", 4),
                "Universi": ("u", 0),
            }

            if row.FONT in font_harfleri:
                if row.FONT != onceki_font:
                    font_harf, adim_sayisi = font_harfleri[row.FONT]
                    pyautogui.click(x=1645, y=233)
                    pyautogui.press('home')
                    pyautogui.press(font_harf)
                    for _ in range(adim_sayisi):
                        pyautogui.press('down')
                    pyautogui.press('enter')
                    pyautogui.click(x=1676, y=383)
                onceki_font = row.FONT

            current_sipno_font_tur = (row.SIPNO, row.FONT, row.TUR)
            if onceki_sipno_font_tur is None or current_sipno_font_tur != onceki_sipno_font_tur:
                pyperclip.copy(row.SIZE)
                pyautogui.click(x=1651, y=143)  # Boyut alani koordinatlari
                pyautogui.click(x=1539, y=201)  # Boyut alani koordinatlari
                pyautogui.click(x=1539, y=201)
                time.sleep(1)
                pyautogui.hotkey('ctrl', 'a')
                pyautogui.hotkey('ctrl', 'v')
                pyautogui.press('enter')
                pyautogui.click(x=1676, y=383)
                pyautogui.press('enter')
                
                pyautogui.click(x=1417, y=146)

            pyautogui.click(x=1676, y=383)
            pyautogui.click(x=85, y=65)

            row.YAZI = noktalari_kaldir(row.YAZI)
            isim = f"{row.SIPNO}-{row.YAZI}-{row.FONT}"
            pyperclip.copy(isim)
            pyautogui.click(x=776, y=487)  # Kayit Yeri Koordinatlari
            pyautogui.click(x=776, y=487)
            time.sleep(1)
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.hotkey('ctrl', 'v')
            pyautogui.press('tab')
            pyautogui.press('t')
            pyautogui.press('enter')
            pyautogui.press('left')
            pyautogui.press('enter')
            pyautogui.press('enter')

            cursor.execute("UPDATE MONOGRAM_DETAY SET YAPILDI = ? WHERE ID = ?", (True, row.ID))
            conn.commit()

            onceki_sipno_font_tur = current_sipno_font_tur

            time.sleep(1)
    conn.close()

def Klasorleme(source_folder, target_root_folder, final_target_folder):
    dosyalar = [f for f in os.listdir(source_folder) if os.path.isfile(os.path.join(source_folder, f))]

    for dosya in dosyalar:
        sipno = dosya.split('-')[0]
        hedef_klasor = os.path.join(target_root_folder, sipno)
        
        if not os.path.exists(hedef_klasor):
            os.makedirs(hedef_klasor)
        
        shutil.move(os.path.join(source_folder, dosya), os.path.join(hedef_klasor, dosya))

    for sipno_klasor in os.listdir(target_root_folder):
        klasor_yolu = os.path.join(target_root_folder, sipno_klasor)
        if os.path.isdir(klasor_yolu):
            for dosya in os.listdir(klasor_yolu):
                dosya_yolu = os.path.join(klasor_yolu, dosya)
                yeni_isim = '-'.join(dosya.split('-')[1:])
                yeni_dosya_yolu = os.path.join(klasor_yolu, yeni_isim)
                os.rename(dosya_yolu, yeni_dosya_yolu)

    for sipno_klasor in os.listdir(target_root_folder):
        kaynak_klasor = os.path.join(target_root_folder, sipno_klasor)
        hedef_klasor = os.path.join(final_target_folder, sipno_klasor)

        if os.path.exists(hedef_klasor):
            shutil.rmtree(hedef_klasor) 

        shutil.move(kaynak_klasor, hedef_klasor)

if __name__ == "__main__":
    Programi_Baslat()
    time.sleep(4)
    SQL_Baglantisi()
    pyautogui.click(x=1700, y=9)
    pyautogui.press('right')
    time.sleep(0.5)
    pyautogui.press('enter')
    print("Islem tamamlandi")
    Klasorleme(source_folder, target_root_folder, final_target_folder)
    source_folder = r"C:\Users\TTBQ-SM\Desktop\0000" 
    target_root_folder = r"\\10.1.10.6\towel\Genel\Monogram\YAPILACAK MONOGRAM\Cizim Bekleyenler" 
    final_target_folder = r"\\10.1.10.6\towel\Genel\Monogram\YAPILACAK MONOGRAM\Hazir Olanlar\Cizilenler"
    print("Dosyalar siparis numarasina gore klasorlere tasindi ve isimler duzenlendi.")
