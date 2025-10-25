import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time

#         _____.___.                .__                         
#         \__  |   |____  _______ __|  |  __  _____                 
#          /   |   \__  \ \___   /  |  | |  |/     \                
#          \____   |/ __ \_/    /|  |  |_|  |  Y Y  \               
#          / ______(____  /_____ \__|____/__|__|_|  /               
#          \/           \/      \/                \/                
#_____.___.                           __________      __            
#\__  |   |____  ___________  ___.__. \____    /____ |  | _______   
# /   |   \__  \ \____ \__  \<   |  |   /     // __ \|  |/ /\__  \  
# \____   |/ __ \|  |_> > __ \\___  |  /     /\  ___/|    <  / __ \_
# / ______(____  /   __(____  / ____| /_______ \___  >__|_ \(____  /
# \/           \/|__|       \/\/              \/   \/     \/     \/ 
#                      ____   ____                                  
#                      \   \ /   /____                              
#                       \   Y   // __ \                             
#                        \     /\  ___/                             
#                         \___/  \___  >                            
#                                    \/                             
#  _____.___.                    __               .__      __       
#  \__  |   |____ ____________ _/  |_ __  ____  __|  |  __|  | __   
#   /   |   \__  \\_  __ \__  \\   __\  |/ ___\|  |  | |  |  |/ /   
#   \____   |/ __ \|  | \// __ \|  | |  \  \___|  |  |_|  |    <    
#   / ______(____  /__|  (____  /__| |__|\___  >__|____/__|__|_ \   
#   \/           \/           \/             \/                \/   

# 47. satıra kullanıcağınız excel dosyasının adını giriniz.
#103. satırda web sitesinde gördüğünüz topluluk seçme dropdownundaki topluluğunuzun sırasını sondan hesaplayarak giriniz. Örneğin kulübünüz listenin sondan 3.sü ise len(options) - 3 olmalı.
#108 - 136. satırlar arasında excel dosyanızdaki sütun isimlerini kodun düzgün çalışması adına belirtiği gibi değiştirin.

print("="*60)
print("Excel Form Gönderme Scripti")
print("="*60)

# BAŞLAŞ SATIRI (hata durumunda kalınan indexi yazarak kodu tekrar çalıştırın.)
START_FROM = 0 

excel_file = "2025uye1.xlsx"    #Excel dosyasını adını buraya giriniz
print(f"\n📄 Excel dosyası okunuyor: {excel_file}")
df = pd.read_excel(excel_file)
df.columns = df.columns.str.strip()

if START_FROM > 0:
    df = df.iloc[START_FROM:]
    print(f"⚠️  {START_FROM + 1}. satırdan devam ediliyor")

print(f"✅ Toplam {len(df)} kayıt işlenecek\n")

chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")

print("🌐 Chrome tarayıcısı başlatılıyor...")
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)
print("✅ Tarayıcı hazır\n")

try:
    url = "http://mediko.akdeniz.edu.tr/topluluk/"
    driver.get(url)
    
    time.sleep(2)
    
    processed_count = 0
    for index, row in df.iterrows():
        print(f"\n[{index + 1}] İşleniyor...")
        
        if processed_count > 0 and processed_count % 50 == 0:
            print("♻️  Tarayıcı yenileniyor...")
            driver.quit()
            time.sleep(3)
            driver = webdriver.Chrome(service=service, options=chrome_options)
            driver.get(url)
            time.sleep(3)
        elif index > START_FROM:
            driver.get(url)
            time.sleep(2)
        
        try:
            wait = WebDriverWait(driver, 15)
            topluluk_select = wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "select"))
            )
            
            select = Select(topluluk_select)
            
            options = select.options
            
            # Sondan üçüncü seçeneği seç (ilk eleman genellikle "Topluluk seçin" olduğu için)
            secilen_topluluk = ""
            if len(options) >= 3:
                # Sondan üçüncü indeks
                secilecek_index = len(options) - 3  #topluluğunuz sırasını sondan hesaplayarak buraya yazın. baştan hesaplamak isterseniz "len(options) - 3" ifadesi yerine direk olarak kulübünüzün sırasından 1 çıkararak yazabilirsiniz. örneğin kulübünüz baştan 15. sırada ise 14 yazınız.
                select.select_by_index(secilecek_index)
                secilen_topluluk = options[secilecek_index].text
                time.sleep(1)
            
            # İsim Soyisim ("İsim soyisim")
            isim_input = driver.find_element(By.CSS_SELECTOR, "input[placeholder='İsim soyisim']")
            isim_input.clear()
            #excel dosyasındaki isim soyisim aldığınız sütünün ismini fullName yapın lütfen
            isim_input.send_keys(row['fullName'])
            
            # Fakülte ("Fakülte ve Bölüm")
            fakulte_input = driver.find_element(By.CSS_SELECTOR, "input[placeholder='Fakülte ve Bölüm']")
            fakulte_input.clear()
            # excel dosyasındaki bölüm/fakülte aldığınız sütünün ismini department yapın lütfen
            fakulte_input.send_keys(row['department'])
            
            # Öğrenci No ("Okul no")
            ogrenci_no_input = driver.find_element(By.CSS_SELECTOR, "input[placeholder='Okul no']")
            ogrenci_no_input.clear()
            # excel dosyasındaki öğrenci numarası aldığınız sütünün ismini studentNumber yapın lütfen
            ogrenci_no_input.send_keys(str(row['studentNumber']))
            
            # Telefon ("Telefon")
            telefon_input = driver.find_element(By.CSS_SELECTOR, "input[placeholder='Telefon']")
            telefon_input.clear()
            # excel dosyasındaki telefon numarası aldığınız sütünün ismini phoneNumber yapın lütfen
            telefon_input.send_keys(str(row['phoneNumber']))
            
            # E-posta ("E-posta")
            eposta_input = driver.find_element(By.CSS_SELECTOR, "input[placeholder='E-posta']")
            eposta_input.clear()
            # excel dosyasındaki e-posta aldığınız sütünün ismini studentEmail yapın lütfen
            email = row['studentEmail'] if pd.notna(row['studentEmail']) else "abc@gmail.com"
            eposta_input.send_keys(email)
            
            # Adres
            adres_input = driver.find_element(By.CSS_SELECTOR, "textarea[placeholder='Adres']")
            adres_input.clear()
            # Sabit adres bilgisi giriliyor
            adres_input.send_keys("Akdeniz Üniversitesi")
            
            submit_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Kaydol')]")
            submit_button.click()
            time.sleep(2)
            
            print(f"✅ {row['fullName']} - {secilen_topluluk} - GÖNDERİLDİ")
            processed_count += 1
            
            time.sleep(2)
            
        except Exception as e:
            print(f"❌ Hata: {str(e)}")
            print(f"⚠️  {index + 1}. satırda hata! Devam ediliyor...")
            try:
                driver.get(url)
                time.sleep(2)
            except:
                pass
            continue
    
    print("\n" + "="*60)
    print("✅ TÜM KAYITLAR İŞLENDI!")
    print("="*60)
    print(f"\nToplam {len(df)} kayıt işlendi.")
    print("\n⚠️  Gerçek gönderim için script içinde yorum satırlarını kaldırın.")
    
except Exception as e:
    print(f"\n❌ Genel hata: {str(e)}")

finally:
    input("\n\n🔴 İşlem tamamlandı. Tarayıcıyı kapatmak için Enter'a basın...")
    driver.quit()


#   \  |           |      |  /                           | 
#  |\/ |   -_)   _| _|    . <    _` |   ` \    _` |   _| |  | 
# _|  _| \___| _| \__|   _|\_\ \__,_| _|_|_| \__,_| _|  _| _| 

                                                                                            