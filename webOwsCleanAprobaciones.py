#Clean approves in OWS SPMS 2

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import yaml, os, time, sys

#Starting time
start_time = time.time()

# Save in desktop
#dest_dir = r"C:\Users\XXXXXXXX\Desktop"
dest_dir = os.path.normpath(os.path.expanduser("~/Desktop"))
options = Options() #object of ChromeOptions
options.add_experimental_option('excludeSwitches', ['enable-logging']) #remove errors in the console
options.add_argument('--headless') # no screen follow up
options.add_argument('--disable-gpu') # no screen follow up
options.add_experimental_option("prefs", {
  "download.default_directory":dest_dir,
  "download.prompt_for_download": False,
  "download.directory_upgrade": True,
  "safebrowsing.enabled": True
})

driver = webdriver.Chrome(options=options)

# Import the ymal file
conf = yaml.full_load(open(r'D:\Python\New FILES\XXXXXXXX\XXXXXXXX.yml'))
myOWSUser = conf['OWSurer']['username']
myOWSPassword = conf['OWSurer']['password']
OWSurl = conf['OWSurer']['url']

#driver = webdriver.Chrome()
driver.maximize_window() # maximize
#driver.minimize_window() # minimize

# Function to log in
def login(url,usernameInput, myOWSUser, password, myOWSPassword, btn_submit):
   driver.get("XXXXXXXX") #SPMS Aprobaciones
   driver.find_element(By.ID,usernameInput).send_keys(myOWSUser)
   driver.find_element(By.ID,password).send_keys(myOWSPassword)
   driver.find_element(By.ID,btn_submit).click()

login(OWSurl, "usernameInput", myOWSUser, "password", myOWSPassword, "btn_submit")
print("We are inside")
driver.implicitly_wait(3)

# Get number of repetitions
#driver.find_element_by_css_selector("[@title^='Eliminar']").click
time.sleep(0.2)
driver.find_element(By.ID,'toolbarSearchButton').click() #search to load the page correctly
totalToClean = driver.find_element(By.ID,'ext-comp-1010').text
print(totalToClean)

# Check if no lines to clean
#if totalToClean == "No records found." or "Ningún registro encontrado.":
if totalToClean == "Ningún registro encontrado.":
  print(f'Not approvals to clean in the page. Exiting!')
  driver.close()
  print(f'[TOTAL]-Done! Completed in {round(time.time()-start_time,2)} seconds.')  
  sys.exit()

totalToCleanNumber = int (''.join(filter(str.isdigit, totalToClean))) #Take only the numbers of the string
print(f'We have to clean: {totalToCleanNumber} lines.')
print(f'Estimated time: {totalToCleanNumber*4.5+7} seconds.')

# Loop to clean all the lines
count = 0
while count < totalToCleanNumber:
  time.sleep(1)
  driver.find_element(By.ID,'toolbarSearchButton').click()
  time.sleep(0.9)
  driver.find_element_by_xpath('//*[@title="Eliminar"]').click()
  time.sleep(0.8)
  driver.find_element(By.ID,'ymPrompt_btn_confirm').click()
  time.sleep(0.8)
  driver.find_element(By.ID,'toolbarSearchButton').click()
  time.sleep(0.8)
  count += 1
  print(f'Line deleted! {count}/{totalToCleanNumber} Continuing...')

# Closing
time.sleep(2)
driver.quit()

print(f'[TOTAL]-Done! Completed in {round(time.time()-start_time,2)} seconds.')
