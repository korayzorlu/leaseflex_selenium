from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select

import time
import pandas as pd
from utils import *

options = Options()

driver = webdriver.Chrome(options=options)
driver.set_window_size(1920, 1080)
wait = WebDriverWait(driver, 10)

driver.get("https://leaseflex.arileasing.com.tr/Ari_Leasing/Logon.aspx")

#login
username_input = driver.find_element(By.ID, "txtUserName")
password_input = driver.find_element(By.ID, "txtPassword")
login_button = driver.find_element(By.ID, "btnLogonEx")

username_input.click()
username_input.send_keys("koray.zorlu")

password_input.click()
password_input.send_keys("Kozo5313-*")

login_button.click()

time.sleep(5)

#go to main frame
# try:
#     driver.switch_to.frame("main")
#     ok_button = driver.find_element(By.CLASS_NAME, "messageDialogOKButton")
#     ok_button.click()
# except:
#     pass


#go to contents frame
driver.switch_to.frame("contents")

#go to cari fiş
genel_muhasebe_button = driver.find_element(By.ID, "ApplicationMenuWebTree_14").find_element(By.CLASS_NAME, "igtr_Root")
genel_muhasebe_button.click()

muhasebe_fisleri_button = driver.find_element(By.ID, "ApplicationMenuWebTree_14_6").find_element(By.CLASS_NAME, "igtr_Parent")
muhasebe_fisleri_button.click()

cari_fisleri_button = driver.find_element(By.ID, "ApplicationMenuWebTree_14_6_3").find_element(By.CLASS_NAME, "igtr_Leaf")
cari_fisleri_button.click()

#go to yeni fiş
driver.switch_to.default_content()
driver.switch_to.frame("main")
time.sleep(0.5)

yeni_fis_button = driver.find_element(By.ID, "btnNewVoucher")
yeni_fis_button.click()
time.sleep(5)

fis_grubu_input = driver.find_element(By.ID, "cmbVoucherGroupTextBoxMain")
fis_grubu_input.click()
fis_grubu_input.send_keys("mahsup")
time.sleep(3)
fis_grubu_option = driver.find_element(By.ID, "Tr0")
fis_grubu_option.click()

fis_tipi_input = driver.find_element(By.ID, "cmbVoucherTypeTextBoxMain")
fis_tipi_input.click()
fis_tipi_input.send_keys("mahsup")
time.sleep(3)
fis_tipi_option = driver.find_element(By.ID, "Tr0")
fis_tipi_option.click()

# add row
add_row_button = driver.find_element(By.ID, "grdLedgerVoucherxGrid_an_0")
add_row_button.click()

# add kira
islem_grubu_block = driver.find_element(By.ID, "grdLedgerVoucherxGrid_rc_0_2").find_element(By.CSS_SELECTOR, "nobr")
islem_grubu_block.click()
islem_grubu_input = driver.find_element(By.ID, "grdLedgerVoucherxGrid_tb")
islem_grubu_input.send_keys("kira")
time.sleep(3)
islem_grubu_option = driver.find_element(By.ID, "Tr0")
islem_grubu_option.click()

#import file
excel_file = pd.ExcelFile("PA-mizan-test.xlsx")
sheet_name = excel_file.sheet_names[2]

file_data = pd.read_excel("PA-mizan-test.xlsx", sheet_name)
df = pd.DataFrame(file_data)

df_sec= df[df["İşlem Grubu"] == "Seç"]
df_sinpas= df[(df["hesap kartno"].str.split('.').str[4] == "001087") & ( df["İşlem Grubu"] != "Seç")]
df = df[(df["hesap kartno"].str.split('.').str[4] != "001087") & ( df["İşlem Grubu"] != "Seç")]

group_size = 150
groups = [df.iloc[i:i + group_size].to_dict(orient='records') for i in range(0, len(df), group_size)]
groups_sec = [df_sec.iloc[i:i + group_size].to_dict(orient='records') for i in range(0, len(df_sec), group_size)]
groups_sinpas = [df_sinpas.iloc[i:i + group_size].to_dict(orient='records') for i in range(0, len(df_sinpas), group_size)]

all_groups = groups + groups_sec + groups_sinpas

all_groups[1][0] = groups_sec[0][0]

# add row
triggered = False
for group in all_groups:
    if not triggered and (group[0]["hesap kartno"].split('.')[4] == "001087" or group[0]["İşlem Grubu"] == "Seç"):
        time.sleep(1)
        
        back_button = driver.find_element(By.ID, "btnBack")
        back_button.click()

        time.sleep(3)

        type_select_element = driver.find_element(By.ID, "ddlAccountType")
        type_select = Select(type_select_element)
        type_select.select_by_visible_text("Banka") 

        #go to yeni fiş
        driver.switch_to.default_content()
        driver.switch_to.frame("main")
        time.sleep(0.5)

        yeni_fis_button = driver.find_element(By.ID, "btnNewVoucher")
        yeni_fis_button.click()
        time.sleep(5)

        fis_grubu_input = driver.find_element(By.ID, "cmbVoucherGroupTextBoxMain")
        fis_grubu_input.click()
        fis_grubu_input.send_keys("mahsup")
        time.sleep(3)
        fis_grubu_option = driver.find_element(By.ID, "Tr0")
        fis_grubu_option.click()

        fis_tipi_input = driver.find_element(By.ID, "cmbVoucherTypeTextBoxMain")
        fis_tipi_input.click()
        fis_tipi_input.send_keys("mahsup")
        time.sleep(3)
        fis_tipi_option = driver.find_element(By.ID, "Tr0")
        fis_tipi_option.click()

        # add row
        add_row_button = driver.find_element(By.ID, "grdLedgerVoucherxGrid_an_0")
        add_row_button.click()

        select_input = driver.find_element(By.ID, f"grdLedgerVoucherxGrid_rc_0_0")
        select_input.click()

        triggered = True
    
    if triggered:
        banka_yontemi(group,driver,By,wait,EC,ActionChains)
    else:
        musteri_yontemi(group,driver,By,wait,EC,ActionChains)

    

    bakiye_block = driver.find_element(By.ID, "igtxtnmVoucherBalance")
    bakiye = bakiye_block.get_attribute("value")

    print(f"bakiye: {bakiye} / tip: {type(bakiye)}")

print("İşlem tamamlandı. Tarayıcıyı kapatmak için pencereyi kendiniz kapatın.")
print("Tarayıcı kapanana kadar bekleniyor...")

try:
    while True:
        # Tarayıcı hala açık mı kontrol et
        driver.title  # bir etkileşim olsun ki hata fırlatsın kapanırsa
except:
    print("Tarayıcı kapatıldı. Program sonlandırılıyor.")


#time.sleep(10)

#driver.quit()