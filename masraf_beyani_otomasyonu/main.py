import pytesseract
import cv2
import sys
import os
import glob
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import *
import re
import pandas as pd

class Pencere(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.dosyaAc = QtWidgets.QPushButton("İşlem Görecek Dosya")
        v_box = QtWidgets.QVBoxLayout()
        v_box.addWidget(self.dosyaAc)
        v_box.addStretch()
        self.setLayout(v_box)
        self.dosyaAc.clicked.connect(self.islemler)
        self.show()

    def islemler(self):
        dosyaAdi = QFileDialog.getExistingDirectory(self, "Klasör Seç", "C:\\")
        dosyaAdi += "/*"
        dosyaYolu = glob.glob(dosyaAdi)
        dosyaYoluFiltre = []
        for i in dosyaYolu:
            if i.endswith("jpg") or i.endswith("jpeg") or i.endswith("png"):
                dosyaYoluFiltre.append(i)

        dosyaYoluKayit = QFileDialog.getSaveFileName(self, "İşlenmiş Dosyaların Kaydedildileceği Yeri Seç",os.getenv("HOME"))
        yeniDosyaYolu = dosyaYoluKayit[0]
        index = yeniDosyaYolu.rfind('/')
        dosyaYoluIsim = yeniDosyaYolu[index + 1::]
        dosyaYoluYol = yeniDosyaYolu[0:index + 1]
        path = os.path.join(dosyaYoluYol, dosyaYoluIsim)
        os.mkdir(path)
        pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
        firmAdi = ""
        fisNo = ""
        topkdv = ""
        toplam = ""
        filtrelenmisIfade=""
        firma_adi_ihtimal = ["a.ş.", "a.s.", "a.s", "ins.san.ve", "ins.san", "as."]
        fis_no_ihtimal = ["fis", "fiş", "fig", "fls", "flg", "lis"]
        kdv_ihtimal = ["topkdv", "toplamkdv"]
        toplam_fiyat_ihtimal = ["toplam", "top"]

        def sayiVarMi(kelimeler):
            for i in kelimeler:
                if i.isdigit():
                    return True
            return False


        def karsilikBulma(ihtimal_dizi, index, text_dizi):
            for l in ihtimal_dizi:
                if l == text_dizi[index]:
                    for x in range(1, 5):
                        if sayiVarMi(text_dizi[index + x]):
                            text_dizi[index + x] = text_dizi[index + x].strip("*")
                            text_dizi[index + x] = text_dizi[index + x].replace(",", ".")
                            return text_dizi[index + x]

        sayac = 0
        firmaIsimStop = 0
        ay_sozluk = {"01": "ocak", "02": "subat", "03": "mart", "04": "nisan", "05": "mayıs", "06": "haziran",
                     "07": "temmuz", "08": "agustos", "09": "eylul", "10": "ekim", "11": "kasim", "12": "aralik",
                     "-1": "AYIKLANAMAYANLAR"}
        for y in dosyaYoluFiltre:
            resim = cv2.imread(y)
            grilestirOnIsle = cv2.cvtColor(resim, cv2.COLOR_BGR2GRAY)
            yazi = pytesseract.image_to_string(grilestirOnIsle)
            yazi = yazi.replace("#", "*")
            yazi = yazi.replace("x", "%")
            yazi = yazi.replace("$", "ş")
            yazi = yazi.lower()
            date_extract_pattern = "(\d{2})\.(\d{2})\.(\d{4})"
            eslesme = re.findall(date_extract_pattern, yazi)
            if len(eslesme) == 0:
                date_extract_pattern = "[0-9]{1,2}\\/[0-9]{1,2}\\/[0-9]{4}"
                eslesme = re.findall(date_extract_pattern, yazi)
            eslesme = str(eslesme)
            print(eslesme)
            tarih = eslesme
            tarih = tarih.strip("[(")
            tarih = tarih.strip(")]")
            tarih = tarih.replace("'", "")
            tarih = tarih.replace(",", "/")
            tarih = tarih.replace(" ", "")
            alinanAy = tarih[3:5]
            try:
                ay_sozluk[alinanAy]
            except KeyError:
                alinanAy = "-1"

            yazi = yazi.split()
            for z in range(0, len(yazi)):
                for l in firma_adi_ihtimal:
                    if l == yazi[z]:
                        for i in reversed(range(0, 3)):
                            firmAdi += yazi[z - i] + " "
                            firmaIsimStop += 1
                    if firmaIsimStop > 1:
                        break

                if karsilikBulma(fis_no_ihtimal, z, yazi) is not None:
                    fisNo = karsilikBulma(fis_no_ihtimal, z, yazi)

                if karsilikBulma(kdv_ihtimal, z, yazi) is not None:
                    topkdv = karsilikBulma(kdv_ihtimal, z, yazi)

                if karsilikBulma(toplam_fiyat_ihtimal, z, yazi) is not None:
                    toplam = karsilikBulma(toplam_fiyat_ihtimal, z, yazi)

            filtre_kelime = [tarih, firmAdi, fisNo, toplam, topkdv]
            for j in range(0, len(filtre_kelime)):
                if filtre_kelime[j] == "":
                    filtre_kelime[j] = "Bulunamadı"
                filtrelenmisIfade = filtre_kelime[0] + "," + filtre_kelime[1] + "," + filtre_kelime[2] + "," + \
                                     filtre_kelime[3] + "," + filtre_kelime[4] + "\n"

            olusturulacakYol = yeniDosyaYolu + "/" + ay_sozluk[alinanAy] + ".csv"
            olusacakDosya = open(olusturulacakYol, "a")

            if len(glob.glob(yeniDosyaYolu + "/*")) == sayac:
                olusacakDosya.write(filtrelenmisIfade)
            else:
                olusacakDosya.write("Tarih,Kimden Alindigi,Fis No,Toplam Tutar,Toplam KDV\n" + filtrelenmisIfade)
                sayac += 1

            firmAdi = ""

        csv_dosya = glob.glob(yeniDosyaYolu + "/*")
        for i in csv_dosya:
            tablo = pd.read_csv(i, encoding='unicode_escape')
            siralanmisTablo=tablo.sort_values(by='Tarih')
            siralanmisTablo['KDV Hariç Tutar']=siralanmisTablo['Toplam Tutar']-siralanmisTablo['Toplam KDV']
            yeniExcelYolu = i[:-3]
            siralanmisTablo.to_excel(yeniExcelYolu + "xlsx", index=None, header=True)

        olusacakDosya.close()

        for h in csv_dosya:
            if os.path.exists(h) and os.path.isfile(h):
                os.remove(h)

app = QtWidgets.QApplication(sys.argv)
pencere = Pencere()
sys.exit(app.exec_())
