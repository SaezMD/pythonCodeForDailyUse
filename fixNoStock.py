from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import yaml, os, time
from datetime import date, timedelta
import pandas as pd

#Starting time
start_time = time.time()

#dest_dir = r"C:\Users\S00492308\Desktop" # Save in desktop
dest_dir = os.path.normpath(os.path.expanduser("~/Desktop"))
options = Options() #object of ChromeOptions
options.add_experimental_option("prefs", {
  "download.default_directory":dest_dir,
  "download.prompt_for_download": False,
  "download.directory_upgrade": True,
  "safebrowsing.enabled": True
})

driver = webdriver.Chrome(chrome_options=options)

#import the yaml file
conf = yaml.full_load(open(r'D:\Python\New FILES\WebLog\configOWS.yml')) 
myOWSUser = conf['OWSurer']['username']
myOWSPassword = conf['OWSurer']['password']
OWSurl = conf['OWSurer']['url']
ControlTicketsUrl = conf['Links']['controlTickets']
SpmsOwsUrl = conf['Links']['SpmsEstados']
SpmsNoStockUrl = conf['Links']['SpmsNoStock']

#driver = webdriver.Chrome()
driver.maximize_window()

#function to log in
def login(url,usernameInput, myOWSUser, password, myOWSPassword, btn_submit):
   driver.get(OWSurl) #Control de Tickets
   driver.find_element(By.ID,usernameInput).send_keys(myOWSUser)
   driver.find_element(By.ID,password).send_keys(myOWSPassword)
   driver.find_element(By.ID,btn_submit).click()

login(OWSurl, "usernameInput", myOWSUser, "password", myOWSPassword, "btn_submit")
print("We are inside")
driver.implicitly_wait(3)

# LOOP
# reading the spreadsheet for cells list
recibidoZonaExcel = pd.read_excel('D:/Python/New FILES/Filtros/RecibidoZonaList.xlsx')
#noStockType = "COMPRADO" # Final status in OWS: CERRADO
# getting the references and the cells
cells = recibidoZonaExcel['ref_interna_item']
cellsSituatuion = recibidoZonaExcel['New Situation'] 
linktest = "https://1041-frapp.teleows.com/app/spl/spms_v2/detalle_pedido_NO_Stock.spl?ref_interna_item="
# iterate through the records
for i in range(len(cells)):
    # Fix No STOCK + new tab
    linkPlusRef = linktest + cells[i]
    driver.execute_script("window.open('"+linkPlusRef+"', 'new_window')")
    driver.switch_to_window(driver.window_handles[-1])
    time.sleep(0.3)
    driver.find_element_by_id('estado_solicitud').clear()
    time.sleep(0.9)
    driver.find_element(By.ID,'estado_solicitud').send_keys(cellsSituatuion[i])
    time.sleep(0.5)
    driver.find_element(By.ID,'submit').click()
    time.sleep(0.3)
    driver.find_element(By.ID,'ymPrompt_btn_confirm').click()
    time.sleep(0.3)
    print(f'{cells[i]} is now in OWS as: {cellsSituatuion[i]}')

# Closing
time.sleep(2)
driver.quit()

print(f'[TOTAL]-Done! Completed in {round(time.time()-start_time,2)} seconds.')
