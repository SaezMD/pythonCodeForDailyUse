from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import yaml, os, time, sys
import pyautogui

#Starting time
start_time = time.time()

# Save in desktop
#dest_dir = r"C:\Users\XXXXXX\Desktop"
dest_dir = os.path.normpath(os.path.expanduser("~/Desktop"))
options = Options() #object of ChromeOptions
options.add_experimental_option('excludeSwitches', ['enable-logging']) #remove errors in the console
#options.add_argument('--headless') # no screen follow up
#options.add_argument('--disable-gpu') # no screen follow up
options.add_experimental_option("prefs", {
  "download.default_directory":dest_dir,
  "download.prompt_for_download": False,
  "download.directory_upgrade": True,
  "safebrowsing.enabled": True
})

driver = webdriver.Chrome(options=options)

# Import the ymal file
conf = yaml.full_load(open(r'D:\Python\New FILES\XXXXXX\XXXXXX.XXXXXX')) 
myOWSUser = conf['OWSurer']['username']
myOWSPassword = conf['OWSurer']['password']
OWSurl = conf['OWSurer']['url']

#driver = webdriver.Chrome()
driver.maximize_window() # maximize
#driver.minimize_window() # minimize

# Function to log in
def login(url,usernameInput, myOWSUser, password, myOWSPassword, btn_submit):
   driver.get("XXXXXX") #SPMS Aprobaciones
   driver.find_element(By.ID,usernameInput).send_keys(myOWSUser)
   driver.find_element(By.ID,password).send_keys(myOWSPassword)
   driver.find_element(By.ID,btn_submit).click()

login(OWSurl, "usernameInput", myOWSUser, "password", myOWSPassword, "btn_submit")
print("We are inside")
driver.implicitly_wait(3)

#Config the screen to only see Cancelled and Pending
driver.find_element(By.ID,'ToolbarExtender1').click() #extend menu

time.sleep(0.3)
driver.find_element(By.ID,'tql_extended').clear()
driver.find_element(By.ID,'tql_extended').send_keys("Ver todos")

driver.find_element(By.ID,'estado_recepcion').send_keys("CANCELADO") #Cancelados+Pendientes
driver.find_element(By.ID,'estado_devolucion').send_keys("PENDIENTE") #Cancelados+Pendientes

#driver.find_element(By.ID,'estado_recepcion').send_keys("ERROR") #Registros con error

driver.find_element(By.ID,'buscar').click() #search button
time.sleep(0.5)

# Get number of repetitions
totalToClean = driver.find_element(By.ID,'ext-comp-1018').text
print(f'Total to clean Cancelados+Pendientes: {totalToClean}')

# Check if no lines to clean for Cancelados+Pendientes
#if totalToClean == "No records found." or "Ningún registro encontrado.":
if totalToClean == "No records found.":
  # If Cancelados+Pendientes are done, continue with Errors to clean:
  time.sleep(0.5)
  driver.find_element(By.ID,'estado_recepcion').clear()
  time.sleep(0.5)
  driver.find_element(By.ID,'estado_devolucion').clear()
  time.sleep(0.5)
  driver.find_element(By.ID,'estado_recepcion').send_keys("ERROR") #Registros con error
  time.sleep(0.5)
  driver.find_element(By.ID,'buscar').click() #search button
  time.sleep(0.5)
  # Get number of repetitions
  totalToClean = driver.find_element(By.ID,'ext-comp-1018').text
  print(f'Total to clean with errors: {totalToClean}')
  time.sleep(0.5)
  
  if totalToClean == "No records found.":
    print(f'Not entries to clean in the page. Exiting!')
    driver.close()
    print(f'[TOTAL]-Done! Completed in {round(time.time()-start_time,2)} seconds.')  
    sys.exit()

totalToCleanNumber = int (''.join(filter(str.isdigit, totalToClean))) #Take only the numbers of the string
print(f'We have to clean: {totalToCleanNumber} lines.')
print(f'Estimated time: {totalToCleanNumber*3.2+10} seconds.')

# Loop to clean all the lines
count = 0
while count < totalToCleanNumber:

  pyautogui.moveTo(860, 600, duration = 0.1)
  pyautogui.hscroll(100) 

  time.sleep(0.6)
  location = pyautogui.locateOnScreen(r'D:\Python\New FILES\imagesByMouse\engranage.png') #Engranaje icon by mouse
  pyautogui.click(location)

  time.sleep(0.6) 
  driver.find_element(By.ID,'aid_9').click() #Cancel line button
  time.sleep(0.6)

  driver.find_element(By.ID,'ymPrompt_btn_confirm').click() #OK Button
  time.sleep(0.6)

  count += 1

# Closing
time.sleep(2)
driver.close()

print(f'[TOTAL]-Done! Completed in {round(time.time()-start_time,2)} seconds.')
