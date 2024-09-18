#Doğrulama Kodu
import requests
from bs4 import BeautifulSoup
url = "https://docs.google.com/spreadsheets/d/1AP9EFAOthh5gsHjBCDHoUMhpef4MSxYg6wBN0ndTcnA/edit#gid=0"
response = requests.get(url)
html_content = response.content
soup = BeautifulSoup(html_content, "html.parser")
first_cell = soup.find("td", {"class": "s2"}).text.strip()
if first_cell != "Aktif":
    exit()
first_cell = soup.find("td", {"class": "s1"}).text.strip()
print(first_cell)


import requests
from bs4 import BeautifulSoup
import pandas as pd
from concurrent.futures import ThreadPoolExecutor


print("Oturum Açma Başarılı Oldu")
print(" /﹋\ ")
print("(҂`_´)")
print("<,︻╦╤─ ҉ - -")
print("/﹋\\")
print("Mustafa ARI")


# Kullanıcı adı ve şifre
username = "mustafa@modaymis.com"
password = "123456"

# Oturum açılacak web sitesi
login_url = "https://www.modaymis.com/kullanici-giris/?ReturnUrl=%2Fadmin"
order_url = "https://www.modaymis.com/FaprikaReturnXls/U4O7G1/1/"

# İstek başlıkları
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36",
    "Referer": "https://www.modaymis.com/",
}

# Oturum açma işlemi
session = requests.Session()
response = session.get(login_url, headers=headers)
soup = BeautifulSoup(response.text, "html.parser")
token = soup.find("input", {"name": "__RequestVerificationToken"}).get("value")
login_data = {
    "EmailOrPhone": username,
    "Password": password,
    "__RequestVerificationToken": token,
}
response = session.post(login_url, data=login_data, headers=headers)

# Excel dosyasını indirme
excel_response = session.get(order_url, headers=headers)

# Excel dosyasını işleme
with open("iadeler.xlsx", "wb") as f:
    f.write(excel_response.content)




import shutil
# Kaynak dosya adı
kaynak_excel = "iadeler.xlsx"

# Kopya dosya adı (istediğiniz adı ve konumu belirtin)
kopya_excel = "iadeler kopya.xlsx"

# Dosyayı kopyala
shutil.copy(kaynak_excel, kopya_excel)



# Kaynak dosya adı
kaynak_excel = "iadeler.xlsx"

# Dosyayı oku
df = pd.read_excel(kaynak_excel)

# "Id" sütunu hariç diğer tüm sütunları sil
df.drop(df.columns.difference(['SiparisId']), axis=1, inplace=True)

# "Id" sütunu hariç diğer tüm sütunları sil
df = df[['SiparisId']].drop_duplicates()

# Sonuçları kaynak Excel dosyasına yaz
df.to_excel(kaynak_excel, index=False)




from tqdm import tqdm


# Okuma işlemi
df = pd.read_excel("iadeler.xlsx")

def get_specific_order_data(order_id):
    order_url = f"https://www.modaymis.com/admin/order/edit/{order_id}/"
    order_response = session.get(order_url, headers=headers)
    order_soup = BeautifulSoup(order_response.text, "html.parser")

    data = {}
    # İstenen verileri al
    for item in order_soup.find_all("div", class_="row static-info align-reverse"):
        key = item.find("div", class_="col-md-8 name").text.strip()
        if key in ["Kargo Bedeli:", "Kapıda Nakit Ödeme Bedeli:", "Vade Farkı:", "Toplam:"]:
            value = item.find("div", class_="col-md-4 value").text.strip()
            data[key] = value



    # "Sipariş Durumu" bilgisini al
    form_group = order_soup.find("div", class_="form-group")
    siparis_durumu = form_group.find("b").text.strip() if form_group and form_group.find("b") else "Bilinmiyor"
    data["Sipariş Durumu"] = siparis_durumu



    # "Kargo Firması" bilgisini al
    label_for_kargo_firmasi = order_soup.find("label", {"for": "ShippingMethodId"})
    if label_for_kargo_firmasi:
        kargo_firmasi = label_for_kargo_firmasi.find_next("div", class_="col-md-9 col-sm-9").text.strip()
        data["Kargo Firması"] = kargo_firmasi
    else:
        data["Kargo Firması"] = "Bilinmiyor"


    





    # "Üçüncü Form Bilgisi"ni al ve birleştir
    form_groups = order_soup.find_all("div", class_="form-group")
    ucuncu_form_bilgisi_etiketler = form_groups[2].find_all("b")
    ucuncu_form_bilgisi = " ".join([b.text.strip() for b in ucuncu_form_bilgisi_etiketler])
    data["Üçüncü Form Bilgisi"] = ucuncu_form_bilgisi

    return data

# "SiparisId" sütunundaki her sipariş için işlem yap (10'arlı gruplar halinde)
with ThreadPoolExecutor(max_workers=50) as executor:
    futures = []
    for order_id in df["SiparisId"]:
        future = executor.submit(get_specific_order_data, order_id)
        futures.append(future)

    # Sadece ikinci döngüde ilerleme çubuğunu göster
    for index, future in enumerate(tqdm(futures, desc="Veriler Çekiliyor", unit="order")):
        order_data = future.result()
        for key, value in order_data.items():
            df.at[index, key] = value

# Excel dosyasını güncelle (mevcut dosyanın üzerine yaz)
with pd.ExcelWriter("iadeler.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    df.to_excel(writer, sheet_name="Returns", index=False)








# Kaynak dosyalar
kaynak_excel = "iadeler.xlsx"
kopya_excel = "iadeler kopya.xlsx"

# Dosyaları oku
df_kaynak = pd.read_excel(kaynak_excel, sheet_name='Returns')
df_kopya = pd.read_excel(kopya_excel)

# "SiparisId" sütununu kullanarak birleştirme işlemi
df_sonuc = pd.merge(df_kopya, df_kaynak[['SiparisId', 'Kargo Bedeli:']], on='SiparisId', how='left')

# "2. Sutun" sütununu "iadeler kopya.xlsx" dosyasındaki ilgili sütuna yaz
df_kopya['Kargo Bedeli:'] = df_sonuc['Kargo Bedeli:']













try:
    df_sonuc = pd.merge(df_kopya, df_kaynak[['SiparisId', 'Kapıda Nakit Ödeme Bedeli:']], on='SiparisId', how='left')

    # "2. Sutun" sütununu "iadeler kopya.xlsx" dosyasındaki ilgili sütuna yaz
    df_kopya['Kapıda Nakit Ödeme Bedeli:'] = df_sonuc['Kapıda Nakit Ödeme Bedeli:']
except Exception as e:
    pass
    pass



df_sonuc = pd.merge(df_kopya, df_kaynak[['SiparisId', 'Toplam:']], on='SiparisId', how='left')

# "2. Sutun" sütununu "iadeler kopya.xlsx" dosyasındaki ilgili sütuna yaz
df_kopya['Toplam:'] = df_sonuc['Toplam:']




df_sonuc = pd.merge(df_kopya, df_kaynak[['SiparisId', 'Sipariş Durumu']], on='SiparisId', how='left')

# "2. Sutun" sütununu "iadeler kopya.xlsx" dosyasındaki ilgili sütuna yaz
df_kopya['Sipariş Durumu'] = df_sonuc['Sipariş Durumu']



try:
    df_sonuc = pd.merge(df_kopya, df_kaynak[['SiparisId', 'Üçüncü Form Bilgisi']], on='SiparisId', how='left')

    # "2. Sutun" sütununu "iadeler kopya.xlsx" dosyasındaki ilgili sütuna yaz
    df_kopya['Üçüncü Form Bilgisi'] = df_sonuc['Üçüncü Form Bilgisi']
except Exception as e:
    pass
    pass



try:
    df_sonuc = pd.merge(df_kopya, df_kaynak[['SiparisId', 'Kargo Firması']], on='SiparisId', how='left')

    # "2. Sutun" sütununu "iadeler kopya.xlsx" dosyasındaki ilgili sütuna yaz
    df_kopya['Kargo Firması'] = df_sonuc['Kargo Firması']

except Exception as e:
    pass
    pass





# Sonucu "iadeler kopya.xlsx" dosyasına yaz
df_kopya.to_excel(kopya_excel, index=False)













import os
# Excel dosyasının adı
excel_dosyasi = "iadeler.xlsx"

# Excel dosyasını sil
try:
    os.remove(excel_dosyasi)
    pass
except FileNotFoundError:
    print(f"{excel_dosyasi} adlı Excel dosyası bulunamadı.")
except Exception as e:
    print(f"Hata oluştu: {e}")

