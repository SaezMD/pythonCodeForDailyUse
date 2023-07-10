#Copy files from Sharepoint
import shutil, time, os, datetime, sys
from sendWebHook import sendTelegram

start_time = time.time()
# Today's date to open the correct file
todayDate = ' ' + datetime.date.today().strftime("%d%m%Y")
print('Today is:' + todayDate)

# Paths and file names
source_dir = 'XXXXX'
#dest_dir = os.path.normpath(os.path.expanduser("~/Desktop"))
dest_dir = os.path.normpath("D:\XXXXX\Spares\XXXXX XXXXX")
inboundFileName = "07 Inbound XXXXX.xlsx"
outboundFileName = "06 Outbound XXXXX.xlsx"

# Check if the Inbound/Outbound files are loaded OK by RICO in the sharepoint
print('Checking if XXXXX has uploaded Inbound/Outbound files')
numberOfFiles = 2
timeToWait = 60 # 1 minute
maxChecks = 15 # 15 minutes

passes = 0
while True:
    checkFiles = 0
    now_time = datetime.datetime.now().strftime("%Y-%m-%d, %H:%M:%S")
    for top, dir, files in os.walk(source_dir,topdown=True):
        for filename in files:
            file_path = os.path.join(top, filename)
            if todayDate in filename and "bound mv" in filename.lower() :  #check dates and stock name in the files
                checkFiles = checkFiles + 1
                #print(f' {file_path} ')  
    passes = passes+1

    if checkFiles == numberOfFiles: #Check if all (2) files are loaded
        print(f'{now_time} --> Inbound/Outbound files are OK loaded by XXXXX. Total loaded and checked: {checkFiles}/{numberOfFiles}. Continuing to copy files...')
        break
    print(f'{now_time} --> Not all the {numberOfFiles} Inbound/Outbound files are loaded in the sharepoint, total loaded for now: {checkFiles}. Wait {timeToWait/60} minutes for {passes}/{maxChecks} checks.')
    sendTelegram (f'{now_time} --> Not all the {numberOfFiles} Inbound/Outbound files are loaded in the sharepoint, total loaded for now: {checkFiles}.')
    time.sleep(timeToWait)
    if passes > maxChecks:
        print(f'{now_time} --> Not all the {numberOfFiles} Inbound/Outbound files are loaded in the sharepoint, total loaded for now: {checkFiles}. Exiting the script.')
        sys.exit()

#Copy the Inbound and Outbound files in server from today to SparesControl:
print(">>> Coping Excel files from date:" + todayDate + " to " + dest_dir + " ...")

for top, dir, files in os.walk(source_dir,topdown=True):
    for filename in files:
        file_path = os.path.join(top, filename)
        #print(f' {filename} ')
        if todayDate in filename and "bound MV" in filename:
            #print(f' {filename} ')
            if "Inbound MV" in filename:
                shutil.copy2(file_path, os.path.join(dest_dir, inboundFileName))
                print(f'>>> Inbound OK. File: {filename} in: {dest_dir}  as {inboundFileName}')
            if "Outbound MV" in filename:
                shutil.copy2(file_path, os.path.join(dest_dir, outboundFileName))
                print(f'>>> Outbound OK. File: {filename} in: {dest_dir} as {outboundFileName}')

print(f'[TOTAL]-Done! Completed in {round(time.time()-start_time,2)} seconds.')
