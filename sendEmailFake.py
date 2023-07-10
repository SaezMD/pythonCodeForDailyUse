#Send the 2 daily emails when there is no new updates

import shutil, time, os, datetime, sys, psutil, ctypes
import pandas as pd
import win32com.client as win32

sys.path.insert(0, "D:\\Python\\New FILES")
from sendWebHook import sendTelegram

#Logger test
import logging
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s : %(levelname)s : %(name)s : %(message)s')
#logging.basicConfig(level=logging.DEBUG) to send to screen
file_handler = logging.FileHandler('logfile.log') #to send to file
file_handler.setFormatter(formatter) #to send to file
logger.addHandler(file_handler) #to send to file

# Logs
"""
logger.debug('Logger is ON')
logger.info('An info message')
logger.warning('Something is not right.')
logger.error('A Major error has happened.')
logger.critical('Fatal error. Cannot continue')
"""

# Check if Outlook is ready:
def is_outlook_running():
    for p in psutil.process_iter(attrs=['pid', 'name']):
        if "OUTLOOK.EXE" in p.info['name']:
            #print("Yes,", p.info['name'], "is running")
            logger.warning("Yes,", p.info['name'], "is running")
            break
    else:
        logger.critical('Outlook is not running. User needs to open it manually.')
        ctypes.windll.user32.MessageBoxW(0, "Open Outlook", "Check Outlook", 1)
        sys.exit()

is_outlook_running()

# Start time milestone:
start_time = time.time()

# Today's date to save the correct file:
todayDate = ' ' + datetime.date.today().strftime("%d%m%Y")
#print('Today is:' + todayDate)
logger.info('Today is:' + todayDate)

# Paths and file names:
source_dir = os.path.normpath(os.path.expanduser("~/Desktop"))
dest_dir = os.path.normpath(os.path.expanduser("~/Desktop"))

stockXXXX = "Stock XXXX XXXX" + todayDate +".xlsx"
stockXXX = "Stock XXXXX XXX" + todayDate +".xlsx"

# Check if the files are present in the desktop from past days:
#print('Checking if files are present in the desktop from past days')
logger.info('Checking if files are present in the desktop from past days')

numberOfFiles = 2
logger.debug(f'Using total files: {numberOfFiles} ')

while True:
    checkFiles = 0
    now_time = datetime.datetime.now().strftime("%Y-%m-%d, %H:%M:%S")
    for top, dir, files in os.walk(source_dir,topdown=True):
        for filename in files:
            file_path = os.path.join(top, filename)
            if "Stock XXXX XXXX " in filename or "Stock XXXXX XXX " in filename:  #check stock email name in the files
                checkFiles = checkFiles + 1
                #print(f' {file_path} ')
                logger.debug(f'Files checked using filter as a filter: Stock XXXX XXXX & Stock XXXXX XXX: {file_path} ')

    if checkFiles == numberOfFiles: #Check if all (2) files are ready in source directory
        #print(f'{now_time} --> Stock email files are OK in {source_dir}. Total loaded and checked: {checkFiles}. Continuing to change file names...')
        logger.info(f'{now_time} --> Stock email files are OK in {source_dir}. Total loaded and checked: {checkFiles}. Continuing to change file names...')
        break
    #print(f'{now_time} --> Not all the {numberOfFiles} Stock email files are present in Desktop, total loaded for now: {checkFiles}. Exiting code!')
    logger.critical(f'{now_time} --> Not all the {numberOfFiles} Stock email files are present in Desktop, total loaded for now: {checkFiles}. Exiting code!')
    sendTelegram (f'{now_time} --> Not all the {numberOfFiles} Stock email files are present in Desktop, total loaded for now: {checkFiles}.')
    #break
    sys.exit()
        
#Coping old Excel files and change date:
#print(">>> Coping old Excel files and change date to:" + todayDate + " ...")
logger.info(">>> Coping old Excel files and change date to:" + todayDate + " ...")

for top, dir, files in os.walk(source_dir,topdown=True):
    for filename in files:
        file_path = os.path.join(top, filename)
        #print(f' {filename} ')
        #logger.debug(f' Files to change: {filename} ')
        if "Stock XXXX XXXX " in filename:
            shutil.copy2(file_path, os.path.join(dest_dir, stockXXXX))
            #print(f'>>> Stock XXXX OK. File: {filename} in: {dest_dir}  as {stockXXXX}')
            logger.info(f'>>> Stock XXXX OK. File: {filename} in: {dest_dir}  as {stockXXXX}')            
        if "Stock XXXXX XXX " in filename:
            shutil.copy2(file_path, os.path.join(dest_dir, stockXXX))
            #print(f'>>> Stock XXXXX XXX OK. File: {filename} in: {dest_dir}  as {stockXXX}')
            logger.info(f'>>> Stock XXXXX XXX OK. File: {filename} in: {dest_dir}  as {stockXXX}')
        
print(f'[Create FAKE files]-Done! Completed in {round(time.time()-start_time,2)} seconds.')
logger.info(f'[Create FAKE files]-Done! Completed in {round(time.time()-start_time,2)} seconds.')

#If there are no files, try to create from OLD stock files

#Attach the files to the emails and send them:

#XXX mail
start_timeEmailXXX = time.time()
# reading the spreadsheet to get the list of emails
email_list = pd.read_excel('D:/Python/New FILES/Filtros/EmailXXX.xlsx')
emails = email_list['Email']
emailforOutlook = "" 

for i in range(len(emails)):
    # for every record get the email addresses
    email = emails[i]
    emailforOutlook = emailforOutlook + ";" + email

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.Subject = 'XXXXX STOCK XXX' + todayDate
mail.BCC = emailforOutlook
mail.HTMLBody = r"""
Dear all,<br><br>
Please, find attached today's stock report.<br><br>
In case you need further details, do not hesitate to contact me.<br><br>
Best regards,<br>
"""
mail.Attachments.Add(os.path.normpath(os.path.join(dest_dir, stockXXX)))

#print(f'[XXX] - Email FAKE ready!')
#mail.Display()
mail.Send()
#print(f'[XXX] - Email FAKE sent!')
logger.warning(f'[XXX] - Email FAKE sent!')
#sendTelegram (f'XXX FAKE email sent.')

print(f'[XXX] - XXX FAKE email Sent! Completed in {round(time.time()-start_timeEmailXXX,2)} seconds.')
logger.info(f'[XXX] - XXX FAKE email Sent! Completed in {round(time.time()-start_timeEmailXXX,2)} seconds.')

#XXXX mail
start_timeEmailXXXX = time.time()
# reading the spreadsheet for email list
email_listXXXX = pd.read_excel('D:/Python/New FILES/Filtros/EmailXXXX.xlsx')
emails = ''
emails = email_listXXXX['Email HW']
emailforOutlook = "" 

for i in range(len(emails)):
    # for every record get the email addresses
    email = emails[i]
    emailforOutlook = emailforOutlook + ";" + email

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.Subject = 'XXXX STOCK XXXXX' + todayDate
mail.BCC = emailforOutlook
mail.HTMLBody = r"""
Dear all,<br><br>
Please, find attached today's stock report for XXXX.<br><br>
In case you need further details, do not hesitate to contact me.<br><br>
Best regards,<br>
"""
mail.Attachments.Add(os.path.normpath(os.path.join(dest_dir, stockXXXX)))

#print(f'[HW Stock] - Email ready!')
#mail.Display()
mail.Send()
#print(f'[HW Stock] - Email sent!')
logger.warning(f'[HW Stock] - Email sent!')
#sendTelegram (f'XXXX FAKE email sent.')
print(f'[HW Stock] - XXXX FAKE email Sent! Completed in {round(time.time()-start_timeEmailXXXX,2)} seconds.')

#Remove new files created
if os.path.exists((os.path.join(dest_dir, stockXXX))):
    os.remove((os.path.join(dest_dir, stockXXX)))
    logger.info(f'File {os.path.join(dest_dir, stockXXX)} deleted.')
if os.path.exists((os.path.join(dest_dir, stockXXXX))):
    os.remove((os.path.normpath(os.path.join(dest_dir, stockXXXX))))
    logger.info(f'File {os.path.join(dest_dir, stockXXXX)} deleted.')

print(f'[ALL] - XXX & XXXX FAKE emails Sent! Completed in {round(time.time()-start_time,2)} seconds.')
sendTelegram (f'XXXX and XXX FAKE emails sent.')
logger.warning(f'[ALL] - XXX & XXXX FAKE emails Sent! Completed in {round(time.time()-start_time,2)} seconds.')

