#pip install selenium
#pip install update
#pip install chromedriver
#pip install pandas
#pip install xlrd
#pip install pypiwin32
#pip install openpyxl

# Add the Chromedriver in the same folder als the python script
download_chromedriver_path = "https://sites.google.com/chromium.org/driver/downloads"
# https://stackoverflow.com/questions/40555930/selenium-chromedriver-executable-needs-to-be-in-path


# Tutorial
# https://medium.com/python-in-plain-english/create-your-browser-automation-robot-with-python-and-selenium-ed0db1d6d65d

# email = ""
# password = ""

# Cases to validate to test the script
# ------------------------------------
# Normal cases
# 1: Case from inbound warehouse to shipment (Streamics + OMS)
# 2: Case with only Streamics promotions (no OMS)
# 3: Case still in production (no able to promote in streamics)
# 4: Case already shipped
# Cancel cases
# 5: Normal order with 2 parts
# x: Normal order with 1 part
# 6: Cancelled order
# 7: Order that is already shipped
# 8: Case stil in production


import time 
import os
import re
# import sys
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from tkinter import *
from tkinter.ttk import *
from tkinter import messagebox
from datetime import datetime
import pandas as pd
import win32com.client
# from selenium.webdriver.common.action_chains import ActionChains
# import openpyxl
# import math

# -----------------------------------------------------------------------------------------------------------

# Define some variables

testenvironment = 1

if testenvironment == 0:
    # Live environment
    streamics_postprocessing_path_general = 'http://leumamsv00001/STREAMICS/PostProcessing'
    oms_portal = 'https://portal.rsprint.com'
    oms_batch_promotion_path = oms_portal + '/Public/CaseManagement/ViewBatchCaseList.aspx'
    streamicsOrderFile_path = os.getcwd() + '/input/20211203_caseIDoderIDFetch.xlsm'
    streamics_scrap_part_base = 1
else:
    # Test environment
    streamics_postprocessing_path_general = 'http://leumamsv00001/STREAMICSV/PostProcessing'
    oms_portal = 'https://rsprintuat.materialise.net/rsprint/Public'
    oms_batch_promotion_path = oms_portal + '/CaseManagement/ViewBatchCaseList.aspx'
    streamicsOrderFile_path = os.getcwd() + '/input/TestEnvironment_caseIDorderIDFetch.xlsm'
    streamics_scrap_part_base = streamics_postprocessing_path_general + '/Part/'
    

oms_path = oms_portal + '/Default.aspx'
streamics_postprocessing_path_order = streamics_postprocessing_path_general + '/Order/'
user_folder = './users'

# Define the status order
# OMS
statusFlowOms = ("Waiting for design parameters","Design","Design rejected","Design QC","Production","Built","Ready to ship","Shipped")
# Streamics
#statusFlowStreamics = ("Inbound warehouse","Incoming cap QC","Built","Sent to subcontractor","Returned from subcontractor","End product QC","Ready to ship in Paal","Shipped to customer")
# Temp for testing purposes
statusFlowStreamics = ('1210 SLS build breakout + sandblasting',
'1231 SLS QC specific',
'1340 SLS color dye',
'1410 SLS QC after surface finishing',
'1950 SLS sent for delivery',
'9410 MOT Inbound warehouse',
'9420 MOT Incoming cap QC (printed part)',
'9430 MOT Built',
'9440 MOT Sent to subcontractor',
'9450 MOT Returned from subcontractor',
'9460 MOT End product QC',
'9470 MOT Ready to ship in Paal',
'Post processing finished')

test = 1

# Link the Streamics status to the OMS status
# The first 5 Streamics statuses are linked to the first OMS status, to prevent that incorrect promotions are done in the OMS.
StreamicsOmsStatusLink = {
  statusFlowStreamics[0]: statusFlowOms[0],
  statusFlowStreamics[1]: statusFlowOms[0],
  statusFlowStreamics[2]: statusFlowOms[0],
  statusFlowStreamics[3]: statusFlowOms[0],
  statusFlowStreamics[4]: statusFlowOms[0],
  statusFlowStreamics[5]: statusFlowOms[4],
  statusFlowStreamics[6]: statusFlowOms[4],
  statusFlowStreamics[7]: statusFlowOms[5],
  statusFlowStreamics[8]: statusFlowOms[5],
  statusFlowStreamics[9]: statusFlowOms[5],
  statusFlowStreamics[10]: statusFlowOms[5],
  statusFlowStreamics[11]: statusFlowOms[6],
  statusFlowStreamics[12]: statusFlowOms[7]
}

# Handles
handle_streamics_postprocessing = 0
handle_oms_view_orders = 1
handle_oms_batch_promotion = 2
handle_oms_order_detail = 3
handle_streamics_scrap_order = 4

email_accounts = ("rkia.elhassani@materialise.be","julie.wellens@materialise.be","kobe.machielsen@materialise.be","laura.janssens@materialise.be" ,"mariska.swolfs@materialise.be" , "sander.van.nieuwenhoven@materialise.be", "pieter-jan.lijnen@materialise.be", "lode.bosmans@materialise.be","flowbuiltproduction@gmail.com")
email_accounts = tuple(sorted(email_accounts))

scrapReasons = ("","","","","","","")

# -----------------------------------------------------------------------------------------------------------


def click_on_view_orders():
    print_with_timestamp(' ')
    print_with_timestamp('Focus on the "view orders" button and click it')
    print_with_timestamp(' ')
    xpathsearch = '//*[@id="ctl00_left_side_ucMm_accM_nwMm"]/ul/li[2]/a'
    button = driver.find_element(By.XPATH, xpathsearch)
    button.click()
    wait_until_element_is_present('xpath','/html/body/form/div[3]/div[3]/div/h1',20)


def click_all_buttons_overview(xpathsearch_firstall,xpathsearch_secondall):
    time.sleep(5)
    print_with_timestamp('Focus on the first "all" button and click it')
    driver.find_element(By.XPATH, xpathsearch_secondall).click()
    print_with_timestamp('Focus on the second "all" button and click it')
    driver.find_element(By.XPATH, xpathsearch_firstall).click()
    print_with_timestamp(' ')
    time.sleep(15)
    if driver.find_element(By.ID, 'ctl00_right_side_openInNewWindow').is_selected() == False:
        driver.find_element(By.ID, "ctl00_right_side_openInNewWindow").click()
        time.sleep(0.5)

def check_exists_by_xpath(xpath):
    try:
        driver.find_element(By.XPATH, xpath)
    except:
        return False
    return True

def check_exists_by_id(id):
    try:
        driver.find_element(By.ID, id)
        print_with_timestamp('   OMS: The element with id ' + id + ' exists. So we need to click the "Take order" button.')
    except:
        print_with_timestamp('   OMS: The element with id ' + id + ' does not exist. So we do not need to click the "Take order" button.')
        return False
    return True

def check_exists_by_type(elementtype, element_id_or_path, ticker):
    result = False
    x = -1
    while result == False and x <= ticker:
        x += 5
        if x > ticker:
            print_with_timestamp('   The script seems to be stuck. Consider restarting the script.')
            x = -1
        try:
            time.sleep(1)
            if elementtype == 'id':
                driver.find_element(By.ID, element_id_or_path)
            elif elementtype == 'xpath':
                driver.find_element(By.XPATH, element_id_or_path)     
            result = True
        except:
            result = False
    return result

def check_exists_by_value(elementtype, element_id_or_path, ticker, textvalue_to_find):
    result = False
    x = 0
    while result == False and x <= ticker:
        x += 1
        if x > ticker:
            print_with_timestamp('   The script seems to be stuck. Consider restarting the script.')
        try:
            time.sleep(1)
            if elementtype == 'id':
                textvalue_found = driver.find_element(By.ID, element_id_or_path).text
            elif elementtype == 'xpath':
                textvalue_found = driver.find_element(By.XPATH, element_id_or_path).text    

            if textvalue_found == textvalue_to_find:
                result = True
            else:
                result = False
        except:
            result = False
    return result

def print_casebycase(listcases):
    if len(listcases) == 0:
        print_with_timestamp('None')
    else:
        for case in listcases:
            print_with_timestamp(case)

def get_current_status():
    time.sleep(1)
    xpathsearch_statusbar = '/html/body/form/div[3]/div[3]/div/div[2]/div[1]/div[4]/div/span'
    current_status_oms = driver.find_element(By.XPATH, xpathsearch_statusbar).text
    if "Cancelled" not in current_status_oms: 
        current_status_oms_index = statusFlowOms.index(current_status_oms)
        print_with_timestamp('   OMS: The current status is: ' + current_status_oms + ' (' + str(current_status_oms_index) + ')')
    else:
        current_status_oms_index = 'X'
        print_with_timestamp('   OMS: The current status is: ' + current_status_oms )
    return current_status_oms, current_status_oms_index

def get_current_status_line(page):
    time.sleep(1)
    if page == 'overview':
        xpathsearch_statusline = '/html/body/form/div[3]/div[3]/div/div/div[2]/div/div[3]/div[3]/div/table/tbody/tr[2]/td[7]/div/div/div[1]/span'
    if page == 'batch':
        xpathsearch_statusline = '/html/body/form/div[3]/div[3]/div/div/div[2]/div[2]/div[3]/div[3]/div/table/tbody/tr[2]/td[9]/div/div/div[1]/span'
    current_status_oms = driver.find_element(By.XPATH, xpathsearch_statusline).text
    if "Cancelled" not in current_status_oms: 
        current_status_oms_index = statusFlowOms.index(current_status_oms)
        print_with_timestamp('   OMS: The current status is: ' + current_status_oms + ' (' + str(current_status_oms_index) + ')')
    else:
        current_status_oms_index = 99999
        print_with_timestamp('   OMS: The current status is: ' + current_status_oms )
    return current_status_oms, current_status_oms_index

def get_production_substatus():
    #time.sleep(10)
    xpathsearch_streamics_status = '/html/body/form/div[3]/div[3]/div/div[2]/div[2]/div/div[6]/div[4]/div[3]/div[3]/div/table/tbody/tr[2]/td[6]'
    wait_until_element_is_present('xpath',xpathsearch_streamics_status,20)
    current_status_oms = driver.find_element(By.XPATH, xpathsearch_streamics_status).text
    print_with_timestamp('   OMS: The current production substatus is: ' + current_status_oms )
    return current_status_oms

def register_case_ID(validation_status,caseid,cc):
    if validation_status == 'valid':
        print_with_timestamp('   OMS: Registering the valid case promotion.')
        oms_caseids_valid.append(caseid)
    elif validation_status == 'invalid':
        print_with_timestamp('   OMS: Registering the invalid case promotion.')
        oms_caseids_invalid.append(caseid)
    else:
        print_with_timestamp(' ')
        print_with_timestamp('Incorrect use of validation function.')
    
    print_with_timestamp(' ')
    print_with_timestamp('These are currently the valid cases:')
    if len(oms_caseids_valid) == 0:
        print_with_timestamp('None')
    else:
        print_with_timestamp(oms_caseids_valid)
    print_with_timestamp(' ')
    print_with_timestamp('These are currently the invalid cases:')
    if len(oms_caseids_invalid) == 0:
        print_with_timestamp('None')
    else:
        print_with_timestamp(oms_caseids_invalid)
    print_with_timestamp(' ')
    print_with_timestamp('These are the remaining cases still to be done:')
    print_with_timestamp(caseids[cc:])
    print_with_timestamp(' ')


def print_with_timestamp(input):
    f=open(logfile, "a+")
    now = datetime.now()
    now_str = str(now.strftime("%Y-%m-%d %H:%M:%S.%f")[:-3])
    if str(type(input)) != "<class 'list'>":
        # Normal print
        toprint = now_str + '    ' + input
        print(toprint)
        # Print for logfile
        f.write(toprint + '\n')
    else:
        # Normal print
        print(now_str + '    ', end = '')
        print(*input, sep = ", ")  
        # Print for logfile
        listcases = ''
        for case in input:
            listcases = listcases + case + ', '
        listcases = listcases[0:len(listcases)-2]
        f.write(str(datetime.now()) + '    ' + listcases + '\n')

def get_streamics_status_info():
    current_status_streamics = driver.find_element(By.CSS_SELECTOR,'.progressBarForOrderedPart .progress-bar-custom .Active').text
    current_status_streamics_index = statusFlowStreamics.index(current_status_streamics)
    return current_status_streamics, current_status_streamics_index

def get_caseids_from_input(raw_input):
    # Get the indices where the case IDs are located
    indices = [m.start() for m in re.finditer('RS2', raw_input)]
    caseids = []
    # For every case, insert it in the case ID array
    for x in indices:
        newcase = raw_input[x:x+12]
        if newcase not in caseids:
            caseids.append(newcase)
    return caseids

def wait_until_element_is_present(xpath_or_id,string_xpath_or_id,time_to_wait):
    element_status = False
    while element_status == False:
        element_status = check_exists_by_type(xpath_or_id,string_xpath_or_id,time_to_wait)

def check_postprocessing_status():
    if driver.find_element(By.XPATH, xpathsearch_postprocessing_finished_part_1).text == '1':
        status = 'Finished'
    if driver.find_element(By.XPATH, xpathsearch_postprocessing_failed_part_1).text == '1':
        status = 'Failed'
    if driver.find_element(By.XPATH, xpathsearch_postprocessing_started_part_1).text == '1':
        status = 'Started'
    return status

# Generate the input screen 
def clicked():
    global raw_input_caseids
    global raw_input_caseids_rebuilt
    global destination_status_streamics
    global destination_status_streamics_index
    global destination_status_oms
    global destination_status_oms_index
    global delay_factor
    global email
    global password
    global user
    raw_input_caseids  = txt.get("1.0","end") # https://www.delftstack.com/howto/python-tkinter/how-to-get-the-input-from-tkinter-text-box/
    raw_input_caseids_rebuilt  = txt_rebuilt.get("1.0","end") # https://www.delftstack.com/howto/python-tkinter/how-to-get-the-input-from-tkinter-text-box/
    # Get the streamics status
    destination_status_streamics = combo.get()
    destination_status_streamics_index = statusFlowStreamics.index(destination_status_streamics)
    # Get the OMS status
    destination_status_oms = StreamicsOmsStatusLink[destination_status_streamics]
    destination_status_oms_index = statusFlowOms.index(destination_status_oms)
    # delay_factor = float(combo_slower.get())
    delay_factor = 1
    # email = entry_email.get()
    email = entry_email.get()
    password = entry_password.get()
    # Get the user that initiated the script
    user = email.split("@")[0]
    user = user.replace('.','_')
    if password == '':
        # Check if there is a file with the password
        user_file_path = user_folder + '/' + user + '.txt'
        pwd_file_present = os.path.exists(user_file_path)
        if pwd_file_present:
            # Open the file
            f = open(user_file_path, 'r')
            password = f.read()
            f.close()
    window.destroy() # Closes the internal loop and lets the script run forward, otherwise it will freeze here.

def print_summary(caseids_summary,caseids_rebuilt_summary,only_invalid):
    print_with_timestamp(' ')
    print_with_timestamp('Scrapped parts')
    print_with_timestamp('--------------')
    if len(caseids_rebuilt_summary) == 0:
        print_with_timestamp('None')
    else:
        if only_invalid == True:
            counter_rebuilts_with_error = 0
        for case in caseids_rebuilt_summary:
            temp = caseids_rebuilt_summary[case]
            printit = True
            if temp != 'Valid':
                printit = True
                if only_invalid == True:
                    counter_rebuilts_with_error += 1
            else:
                if only_invalid == True:
                    printit = False
            if printit:
                print_with_timestamp(case + ': ' + caseids_rebuilt_summary[case])
        if only_invalid == True:
            if counter_rebuilts_with_error == 0:
                print_with_timestamp('None')
    print_with_timestamp(' ')
    print_with_timestamp('Promotions         Streamics   +   OMS')
    print_with_timestamp('------------------------------------------')
    if len(caseids_summary) == 0:
        print_with_timestamp('None')
    else:
        if only_invalid == True:
            counter_promotions_with_error = 0
        for case in caseids_summary:
            len_streamics = len(caseids_summary[case]['Streamics'])
            temp1 = caseids_summary[case]['Streamics']
            temp2 = caseids_summary[case]['OMS']
            print_spacer = (16 - len_streamics) * ' '
            printit = True
            if temp1 != 'Valid' or temp2 != 'Valid':
                printit = True
                if only_invalid == True:
                    counter_promotions_with_error += 1
            else:
                if only_invalid == True:
                    printit = False
            if printit:
                print_with_timestamp(case + ':      ' + caseids_summary[case]['Streamics'] + print_spacer + caseids_summary[case]['OMS'])
        if only_invalid == True:
            if counter_promotions_with_error == 0:
                print_with_timestamp('None')


# Prepare all necessary stuff to print all output to file
directory = './log'
if not os.path.exists(directory):
    os.makedirs(directory)
datepieces = dict()
datepieces['y'] = str(datetime.now().year)
datepieces['mo'] = str(datetime.now().month)
datepieces['d'] = str(datetime.now().day)
datepieces['h'] = str(datetime.now().hour)
datepieces['mi'] = str(datetime.now().minute)
datepieces['s'] = str(datetime.now().second)
for key in datepieces:
    if len(datepieces[key]) == 1:
        datepieces[key] = '0' + datepieces[key]


# Define the width of the input fields
width_inputfield = 40

window = Tk()
window.title("Status promotions")
window.geometry('800x500')

# Col 0 and 1

label_spacer0 = Label(window, text="          ")
label_spacer0.grid(column=0, row=0)

label_spacer2 = Label(window, text=" ")
label_spacer2.grid(column=1, row=0)

label_input = Label(window, text="Insert the case IDs for normal promotion:")
label_input.grid(column=1, row=1)
txt = Text(window,width=width_inputfield,height=10)
txt.grid(column=1, row=2)

label_spacer3 = Label(window, text=" ")
label_spacer3.grid(column=1, row=3)

label_combo = Label(window,text="Choose a status:")
label_combo.grid(column=1, row=4)
combo = Combobox(window,width=width_inputfield,height=len(StreamicsOmsStatusLink)+1)
combo['values']= ("Choose a status",) + statusFlowStreamics
combo.current(0) #set the selected item
combo.grid(column=1, row=5)

label_spacer = Label(window, text=" ")
label_spacer.grid(column=1, row=6)

Livit = IntVar()
Checkbutton(window, text="Livit cases?", variable=Livit).grid(column=1, row=7)

streamics = IntVar(value=1)
Checkbutton(window, text="Promote in Streamics?", variable=streamics).grid(column=1, row=8)

label_spacer4 = Label(window, text=" ")
label_spacer4.grid(column=1, row=11)

label_email = Label(window, text="Email address")
label_email.grid(column=1, row=12)

entry_email = Combobox(window,width=width_inputfield)
entry_email['values']= ("Choose an email address",) + email_accounts

entry_email.current(0) #set the selected item
entry_email.grid(column=1, row=13)

label_spacer5 = Label(window, text=" ")
label_spacer5.grid(column=1, row=14)

label_password = Label(window,text="Password")
label_password.grid(column=1, row=15)

entry_password = Entry(window,show="*",width=width_inputfield,text="")
entry_password.grid(column=1, row=16)

label_spacer6 = Label(window, text=" ")
label_spacer6.grid(column=1, row=17)

btn = Button(window, text="Start promotion", command=clicked)
btn.grid(column=1, row=18)

# Col 2 and 3

label_spacer0 = Label(window, text="       ")
label_spacer0.grid(column=2, row=1)

label_input_rebuilt = Label(window, text="Insert the case IDs for rebuilt + cancel parts:")
label_input_rebuilt.grid(column=3, row=1)
txt_rebuilt = Text(window,width=width_inputfield,height=10)
txt_rebuilt.grid(column=3, row=2)

# label_spacer3 = Label(window, text=" ")
# label_spacer3.grid(column=3, row=3)

# label_combo = Label(window,text="Choose a scrap reason:")
# label_combo.grid(column=3, row=4)
# combo_scrap = Combobox(window,width=width_inputfield,height=len(StreamicsOmsStatusLink)+1)
# combo_scrap['values']= ("Choose a status",) + scrapReasons
# combo_scrap.current(0) #set the selected item
# combo_scrap.grid(column=3, row=5)

# label_spacer = Label(window, text=" ")
# label_spacer.grid(column=3, row=6)

window.mainloop()
# After the click on the button, the window is destroyed, so data can not be collected again. Check function 'clicked'.

logfile = 'log/' + datepieces['y'] + datepieces['mo'] + datepieces['d'] + '_' + datepieces['h'] + datepieces['mi'] + datepieces['s'] + '_' + user + '_logfile.txt'


# Get the cases for rebuilt and for promotion
caseids = get_caseids_from_input(raw_input_caseids)
#caseids_summary = {}
caseids_rebuilt = get_caseids_from_input(raw_input_caseids_rebuilt)
caseids_rebuilt_summary = {}

# See if it is for Livit or not
if Livit.get() == 1:
    shouldBeLivit = True
else:
    shouldBeLivit = False

if streamics.get() == 1:
    # If the checkbox is checked, keep the Streamics steps
    promote_in_streamics = True
else:
    # If the streamcis checkbox is disabled is disabled
    if len(caseids_rebuilt) == 0:
        # If there are no rebeuilt, don't open streamics
        promote_in_streamics = False
    else:
        # If there are rebuilts, do open streamics
        promote_in_streamics = True


# Check if the Excel file with the link between case ID and Streamics order ID is present.
if os.path.exists(streamicsOrderFile_path) and promote_in_streamics:
    # Start an instance of Excel
    xlapp = win32com.client.DispatchEx("Excel.Application")
    # Open the workbook in said instance of Excel
    wb = xlapp.workbooks.open(streamicsOrderFile_path)
    # Optional, e.g. if you want to debug
    #xlapp.Visible = True
    # Refresh all data connections.
    wb.RefreshAll()
    xlapp.CalculateUntilAsyncQueriesDone()
    #wb.Save()
    wb.Close(SaveChanges=1)
    # Quit
    xlapp.Quit()
    # Now read the refreshed data
    df = pd.read_excel(streamicsOrderFile_path)
    # Create a dict with the link between case ID and streamics order ID
    streamics_order_ids = {} 
    for index, row in df.iterrows():
        ccid = row['Unnamed: 0']
        if isinstance(ccid, str) == True:
            if ccid[0:2] == 'RS':
                coid = row['Unnamed: 1'].split('_')[0]
                if type(ccid) is str and type(coid) is str:
                    print(ccid + '   ' + coid)
                    streamics_order_ids[ccid] = coid



print_with_timestamp('User: ' + user)
print_with_timestamp('Destination status Streamics: ' + destination_status_streamics)
print_with_timestamp('Destination status OMS: ' + destination_status_oms)
print_with_timestamp(' ')
print_with_timestamp('The cases to be processed for promotion:')
print_with_timestamp(caseids)
print_with_timestamp(' ')
print_with_timestamp('The cases to be processed for rebuilt:')
print_with_timestamp(caseids_rebuilt)
print_with_timestamp(' ')

if len(caseids) > 0 or len(caseids_rebuilt) > 0:
    # Startup the browser
    print_with_timestamp('Setting up the browser and checking compatability')
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging']) #https://stackoverflow.com/questions/64927909/failed-to-read-descriptor-from-node-connection-a-device-attached-to-the-system
    service = Service("./drivers/chromedriver.exe")
    driver = webdriver.Chrome(service=service,options=chrome_options)

    # Check the version of the chromedriver
    chrome_browser_version = driver.capabilities['browserVersion'].split('.')[0]
    chromedriver_version = driver.capabilities['chrome']['chromedriverVersion'].split('.')[0]
    print_with_timestamp('Chrome browser version: ' + chrome_browser_version)
    print_with_timestamp('Chromedriver version: ' + chromedriver_version)
    if chrome_browser_version != chromedriver_version: 
        print_with_timestamp('Chrome browser version (' + chrome_browser_version + ') and Chromedriver version (' + chromedriver_version + ') do not match. Please update the Chromedriver!' )
        print_with_timestamp('Please download this from the following url and select version ' + chrome_browser_version + '.x.xxxx.xx.')
        print_with_timestamp(download_chromedriver_path)
    else:
        print_with_timestamp('Versions match')

    if promote_in_streamics == True:
        # Open Streamics
        driver.switch_to.window(driver.window_handles[handle_streamics_postprocessing])
        if testenvironment == 0:
            sdfsf = 1
        else:
            driver.get(streamics_postprocessing_path_general)
        
        # Login to Streamics
        # Insert credentials if available
        if email != "":
            driver.find_element(By.ID, "UserName").clear()
            # Grab the login field
            searchinput = driver.find_element(By.ID, "UserName")
            # Insert the email address
            searchinput.send_keys(email)
        if password != "":
            driver.find_element(By.ID, "Password").clear()
            # Grab the password field
            searchinput = driver.find_element(By.ID, "Password")
            # Insert the password 
            searchinput.send_keys(password)
        if password != "":
            # Press the button
            button = driver.find_element(By.ID, "login-button")
            button.click()
        if password == "":
            # Give some time to insert the credentials and press enter
            time.sleep(20)
        xpathsearch_welcome_streamics = '/html/body/div[1]/div[2]/div/div/div/h1'
        wait_until_element_is_present('xpath',xpathsearch_welcome_streamics,20)


    # Open the OMS
    driver.execute_script("window.open('');")
    driver.switch_to.window(driver.window_handles[handle_oms_view_orders])
    driver.get(oms_path)
    # Apply the delay factor twice to the regular timeout, so only the most critical steps are enforced squared
    driver.set_page_load_timeout(500)

    # Insert credentials if available
    if email != "":
        driver.find_element(By.ID, "ctl00_left_side_ucL_loginView_accP_vL_cL_UserName").clear()
        # Grab the login field
        searchinput = driver.find_element(By.ID, "ctl00_left_side_ucL_loginView_accP_vL_cL_UserName")
        # Insert the email address
        searchinput.send_keys(email)
    if password != "":
        driver.find_element(By.ID, "ctl00_left_side_ucL_loginView_accP_vL_cL_Password").clear()
        # Grab the password field
        searchinput = driver.find_element(By.ID, "ctl00_left_side_ucL_loginView_accP_vL_cL_Password")
        # Insert the password 
        searchinput.send_keys(password)
    if password != "":
        # Press the button
        button = driver.find_element(By.ID, "ctl00_left_side_ucL_loginView_accP_vL_cL_LoginButton")
        button.click()
    if password == "":
        # Give some time to insert the credentials and press enter
        time.sleep(20)

    xpath_vieworders = '/html/body/form/div[3]/div[2]/div[1]/div/div[1]/ul/li[2]/a'
    wait_until_element_is_present('xpath',xpath_vieworders,20)
    
    # Get the OMS role
    xpathsearch_OmsRole = 'ctl00_left_side_ucL_loginView_accL_vLi_lblR'
    element_status = False
    while element_status == False:
        element_status = check_exists_by_type('id',xpathsearch_OmsRole,20)
    current_role = driver.find_element(By.ID, xpathsearch_OmsRole).text
    if current_role == 'Mat Admin':
        xpathsearch_secondall = '//*[@id="ctl00_right_side_RadioButtonOwnerFilter"]/label[2]/span'
        xpathsearch_secondall_pressed = '//*[@id="ctl00_right_side_RadioButtonOwnerFilter"]/label[2]'
    else:
        xpathsearch_secondall = '//*[@id="ctl00_right_side_RadioButtonOwnerFilter"]/label[3]/span'
        xpathsearch_secondall_pressed = '//*[@id="ctl00_right_side_RadioButtonOwnerFilter"]/label[3]'
            

    # Focus on the 'view orders' button and click it
    click_on_view_orders()

    # Focus on the first 'all' button and click it
    # print_with_timestamp('Focus on the first "all" button and click it')
    xpathsearch_firstall = '//*[@id="ctl00_right_side_RadioButtonStatusFilter"]/label[4]/span'
    xpathsearch_firstall_pressed = '//*[@id="ctl00_right_side_RadioButtonStatusFilter"]/label[4]'
    wait_until_element_is_present('xpath',xpathsearch_firstall,20)
    click_all_buttons_overview(xpathsearch_firstall,xpathsearch_secondall)
    #time.sleep(5)
    # driver.find_element(By.XPATH, xpathsearch_firstall).click()

    # Focus on the second 'all' button and click it
    #print_with_timestamp('Focus on the second "all" button and click it')
    # driver.find_element(By.XPATH, xpathsearch_secondall).click()

    # click_all_buttons_overview(xpathsearch_firstall,xpathsearch_secondall)

    # Check the box to open every case in a new tab
    # driver.find_element(By.ID, "ctl00_right_side_openInNewWindow").click()



    # Open a second tab with the batch promotions
    print_with_timestamp('Opening batch promotion in a new tab')
    driver.execute_script("window.open('');")
    driver.switch_to.window(driver.window_handles[handle_oms_batch_promotion])
    time.sleep(0.1)
    driver.get(oms_batch_promotion_path)
    #time.sleep(15)

    # Click the batch promotion all button
    print_with_timestamp('Focus on the batch promotion "all" button and click it')
    print_with_timestamp(' ')
    xpathsearch_batchall = '/html/body/form/div[3]/div[3]/div/div/div[1]/div/span/label[2]/span'
    wait_until_element_is_present('xpath',xpathsearch_batchall,20)  
    time.sleep(1)
    driver.find_element(By.XPATH, xpathsearch_batchall).click()
    time.sleep(3)


    xpathsearch_overview_parts = '/html/body/div[1]/div[2]/div/div/div[3]/div/div/div[1]/span'
    global xpathsearch_postprocessing_started_part_1
    global xpathsearch_postprocessing_started_part_2
    global xpathsearch_postprocessing_finished_part_1
    global xpathsearch_postprocessing_finished_part_2
    global xpathsearch_postprocessing_failed_part_1
    global xpathsearch_postprocessing_failed_part_2
    xpathsearch_postprocessing_started_part_1 = '/html/body/div[1]/div[2]/div/div/div[3]/div/div/div[2]/table[1]/tbody/tr[3]/td[7]'
    xpathsearch_postprocessing_started_part_2 = '/html/body/div[1]/div[2]/div/div/div[3]/div/div/div[2]/table[2]/tbody/tr[3]/td[7]'
    xpathsearch_postprocessing_finished_part_1 = '/html/body/div[1]/div[2]/div/div/div[3]/div/div/div[2]/table[1]/tbody/tr[3]/td[8]'
    xpathsearch_postprocessing_finished_part_2 = '/html/body/div[1]/div[2]/div/div/div[3]/div/div/div[2]/table[2]/tbody/tr[3]/td[8]'
    xpathsearch_postprocessing_failed_part_1 = '/html/body/div[1]/div[2]/div/div/div[3]/div/div/div[2]/table[1]/tbody/tr[3]/td[9]'
    xpathsearch_postprocessing_failed_part_2 = '/html/body/div[1]/div[2]/div/div/div[3]/div/div/div[2]/table[2]/tbody/tr[3]/td[9]'
    xpath_no_parts_to_process = '/html/body/div[1]/div[2]/div/div/div[4]/div/span'
    xpathsearch_expand_streamics_card = '/html/body/div[1]/div[2]/div/div/div[4]/div/div/div[1]'
    xpathsearch_plus_sign_1_scrap = '/html/body/div[1]/div[2]/div/div/div[4]/div/div/div[2]/table/tbody/tr[1]/td[7]/a/em'
    xpathsearch_physical_part_id_1 = '/html/body/div[1]/div[2]/div/div/div[4]/div/div/div[2]/table/tbody/tr[3]/td/div/div[1]/table/tbody/tr/td[1]/a'
    xpathsearch_plus_sign_2_scrap = '/html/body/div[1]/div[2]/div/div/div[4]/div/div/div[2]/table/tbody/tr[4]/td[7]/a/em'
    xpathsearch_scrap_1 = '/html/body/div[1]/div[2]/div[1]/div/div[2]/div/div/div[1]/div[2]/form/a[2]'
    xpathsearch_reason_scrap_1 = '/html/body/div[5]/div/div/div/div/div[2]/div[2]/form/div[1]/div[2]/h3[12]'
    xpathsearch_reason = '/html/body/div[5]/div/div/div/div/div[2]/div[2]/form/div[1]/div[2]/div[12]/p[3]/label/span'
    xpathsearch_scrap_button = 'btnScrap'
    xpathsearch_scrapped_confirmation = '/html/body/div[1]/div[2]/div[1]/div/div[2]/div/div/div[2]/div[1]'

    if promote_in_streamics == True:
        if len(caseids_rebuilt) > 0:
            print_with_timestamp(' ')
            if len(caseids_rebuilt) == 1:
                pairs = 'pair'
            else:
                pairs = 'pairs'
            print_with_timestamp('First things first: scrapping insoles (' + str(len(caseids_rebuilt)) + ' ' + pairs + ')')
            # Go over all case IDs for rebuilt and scrap them
            for caseid_rebuilt in caseids_rebuilt:
                orderid = streamics_order_ids[caseid_rebuilt]
                driver.switch_to.window(driver.window_handles[handle_streamics_postprocessing])
                driver.get(streamics_postprocessing_path_order + orderid)
                print_with_timestamp(' ')
                print_with_timestamp('Scrapping: ' + caseid_rebuilt + ' (' + orderid + ') started')
                time.sleep(0.2)
                # Check if the number exist and that you don't end up with the default page.
                if driver.current_url == streamics_postprocessing_path_order + orderid:
                    # Check if the part is not failed already
                    wait_until_element_is_present('xpath',xpathsearch_overview_parts,20)
                    driver.find_element(By.XPATH, xpathsearch_overview_parts).click()
                    wait_until_element_is_present('xpath',xpathsearch_postprocessing_failed_part_1,20)
                    postprocessing_status = check_postprocessing_status()
                    if postprocessing_status == 'Started':
                        # Get the Streamics tap
                        driver.switch_to.window(driver.window_handles[handle_streamics_postprocessing])
                        driver.get(streamics_postprocessing_path_order + orderid)
                        wait_until_element_is_present('xpath',xpathsearch_overview_parts,20)
                        time.sleep(0.2)
                        if check_exists_by_xpath(xpathsearch_expand_streamics_card):
                            # Check how many parts are still in the order
                            if check_exists_by_xpath(xpathsearch_plus_sign_2_scrap):
                                number_of_parts = 2
                                print_with_timestamp('   STREAMICS: There are ' + str(number_of_parts) + ' parts in the order.')
                            else:
                                number_of_parts = 1
                                print_with_timestamp('   STREAMICS: There is ' + str(number_of_parts) + ' part in the order.')
                            for side in range(1,number_of_parts+1):
                                # Open the parts of the order (expand card)
                                wait_until_element_is_present('xpath',xpathsearch_expand_streamics_card,20)
                                driver.find_element(By.XPATH, xpathsearch_expand_streamics_card).click()
                                # Click the plus sign
                                wait_until_element_is_present('xpath',xpathsearch_plus_sign_1_scrap,20)
                                driver.find_element(By.XPATH, xpathsearch_plus_sign_1_scrap).click()
                                # Click the part number
                                wait_until_element_is_present('xpath',xpathsearch_physical_part_id_1,20)
                                driver.find_element(By.XPATH, xpathsearch_physical_part_id_1).click()
                                # Click the scrap button
                                wait_until_element_is_present('xpath',xpathsearch_scrap_1,20)
                                driver.find_element(By.XPATH, xpathsearch_scrap_1).click()
                                # Click the motion specific text
                                wait_until_element_is_present('xpath',xpathsearch_reason_scrap_1,20)
                                driver.find_element(By.XPATH, xpathsearch_reason_scrap_1).click()
                                # Click the reason - in this case: colateral damage
                                wait_until_element_is_present('xpath',xpathsearch_reason,20)
                                driver.find_element(By.XPATH, xpathsearch_reason).click()
                                # Press the confirm button
                                wait_until_element_is_present('id',xpathsearch_scrap_button,20)
                                driver.find_element(By.ID,xpathsearch_scrap_button).click()
                                # Verify if the part is scrapped
                                wait_until_element_is_present('xpath',xpathsearch_scrapped_confirmation,20)
                                driver.find_element(By.XPATH, xpathsearch_scrapped_confirmation).text
                                print_with_timestamp('   STREAMICS: Scrapping: ' + caseid_rebuilt + ' = ' + str(side) + '/' + str(number_of_parts) + ' succesfull')
                                # Go back to the previous page
                                time.sleep(2)
                                driver.get(streamics_postprocessing_path_order + orderid)
                            print_with_timestamp(' ')
                            caseids_rebuilt_summary[caseid_rebuilt] = 'Valid'
                        else:
                            message = 'Nothing to scrap, still in production (' + orderid + ')'
                            print_with_timestamp('   STREAMICS: ' + message)
                            print_with_timestamp(' ')
                            caseids_rebuilt_summary[caseid_rebuilt] = message
                    else:
                        if postprocessing_status == 'Failed':
                            message = 'The parts are already scrapped (' + orderid + ')'
                            print_with_timestamp('   STREAMICS: ' + message)
                            caseids_rebuilt_summary[caseid_rebuilt] = message
                        if postprocessing_status == 'Finished':
                            message = "The parts are in status 'Post processing finished' and can't be scrapped (" + orderid + ")"
                            print_with_timestamp('   STREAMICS: ' + message)
                            caseids_rebuilt_summary[caseid_rebuilt] = message
                else:
                    message = "The order ID does not exist, it can't be scrapped (" + orderid + ")"
                    print_with_timestamp('   STREAMICS: ' + message)
                    caseids_rebuilt_summary[caseid_rebuilt] = message

    # Go over all case IDs for promotion
    cc = 0
    oms_caseids_valid = []
    oms_caseids_invalid = []
    timestamps = []
    caseids_summary = {}
    for caseid in caseids:
        # Register the timestamp and other vars
        timestamps.append(datetime.now)
        no_error = True
        cc += 1

        caseids_summary[caseid] = {}
        caseids_summary[caseid]['Streamics'] = '-'
        caseids_summary[caseid]['OMS'] = '-'

        print_with_timestamp(' ')
        print_with_timestamp('Processing: ' + caseid + ' (' + str(cc) + '/' + str(len(caseids)) + ')')

        if promote_in_streamics:
            # Get the order ID for this case ID
            orderid = streamics_order_ids[caseid]

            # STREAMICS
            # ---------

            driver.switch_to.window(driver.window_handles[handle_streamics_postprocessing])
            driver.get(streamics_postprocessing_path_order + orderid)
            time.sleep(0.1)

            # Check if the part is not failed already
            wait_until_element_is_present('xpath',xpathsearch_overview_parts,20)
            driver.find_element(By.XPATH, xpathsearch_overview_parts).click()
            wait_until_element_is_present('xpath',xpathsearch_postprocessing_failed_part_1,20)
            postprocessing_status = check_postprocessing_status()
            if postprocessing_status == 'Started':
                # Close the overview again
                driver.find_element(By.XPATH, xpathsearch_overview_parts).click()
                time.sleep(0.2)
                # Check if the card is present
                if check_exists_by_xpath(xpathsearch_expand_streamics_card):
                    # Open the parts of the order (expand card)
                    driver.find_element(By.XPATH, xpathsearch_expand_streamics_card).click()
                    # Get the Streamics status of the parts
                    current_status_streamics, current_status_streamics_index = get_streamics_status_info()
                    print_with_timestamp('   STREAMICS: Destination status is "' + destination_status_streamics + '"')
                    # Compare the current status with the destination status, if smaller, promote to the next step
                    if current_status_streamics_index < destination_status_streamics_index:
                        print_with_timestamp('   STREAMICS: Starting promotions')
                        while current_status_streamics_index < destination_status_streamics_index:
                            # Click the button
                            driver.find_element(By.ID, 'completeAllButton').click()
                            #wait_until_element_is_present('xpath',xpathsearch_expand_streamics_card,20)
                            #time.sleep(1)
                            # Refresh the current status vars
                            if current_status_streamics_index == destination_status_streamics_index-1:
                                # Exception for last status
                                current_status_streamics = 'Post processing finished'
                                current_status_streamics_index = len(StreamicsOmsStatusLink)-1
                            else:
                                # Normal flow
                                wait_until_element_is_present('xpath',xpathsearch_expand_streamics_card,20)
                                time.sleep(1)
                                driver.find_element(By.XPATH, xpathsearch_expand_streamics_card).click()
                                current_status_streamics, current_status_streamics_index = get_streamics_status_info()
                            print_with_timestamp('   STREAMICS: Promoted to "' + current_status_streamics + '"')
                            caseids_summary[caseid]['Streamics'] = 'Valid'
                    else:
                        print_with_timestamp('   STREAMICS: Nothing to promote for ' + caseid + ' (' + orderid + ')')
                        caseids_summary[caseid]['Streamics'] = 'Nothing to promote'
                else:
                    print_with_timestamp('   STREAMICS: Nothing to promote for ' + caseid + ' (' + orderid + ')')
                    caseids_summary[caseid]['Streamics'] = 'Nothing to promote'
            else:
                if postprocessing_status == 'Failed':
                    print_with_timestamp('   STREAMICS: The parts are scrapped at this point.')
                    caseids_summary[caseid]['Streamics'] = 'Parts are scrapped'
                if postprocessing_status == 'Finished':
                    print_with_timestamp('   STREAMICS: The parts are already in status "Post processing finished".')
                    caseids_summary[caseid]['Streamics'] = 'Parts are finished already'
        else:
            caseids_summary[caseid]['Streamics'] = 'Valid'


        # OMS
        # ---

        # Go back to the view orders screen
        driver.switch_to.window(driver.window_handles[handle_oms_view_orders])
        time.sleep(0.1)

        #try:
        # See if the all buttons are still correct
        check_first_all = driver.find_element(By.XPATH, xpathsearch_firstall_pressed).get_attribute('aria-pressed')
        check_second_all = driver.find_element(By.XPATH, xpathsearch_secondall_pressed).get_attribute('aria-pressed')
        if check_first_all != 'true' or check_second_all != 'true':
            click_all_buttons_overview(xpathsearch_firstall,xpathsearch_secondall)
         
        
        # Clear the case ID search field
        wait_until_element_is_present('id','gs_CaseID',20)
        driver.find_element(By.ID, "gs_CaseID").clear()
        # Grab the case ID search field
        searchinput = driver.find_element(By.ID, "gs_CaseID")
        # Insert the case ID into the case ID search field
        searchinput.send_keys(caseid)
        time.sleep(1)
        searchinput.send_keys(Keys.ENTER)
        # Define the Xpath to the case ID in the result table
        xpathsearch_id = '/html/body/form/div[3]/div[3]/div/div/div[2]/div/div[3]/div[3]/div/table/tbody/tr[2]'
        xpathsearch_caseid = '/html/body/form/div[3]/div[3]/div/div/div[2]/div/div[3]/div[3]/div/table/tbody/tr[2]/td[2]'
        wait_until_element_is_present('xpath',xpathsearch_caseid,20)
        # element_status = False
        # while element_status == False:
        #     element_status = check_exists_by_value('xpath', xpathsearch_caseid, 30, caseid)
        # Grab the case ID in the result table row
        caseid_nr = driver.find_element(By.XPATH, xpathsearch_id).get_attribute("id")
        print_with_timestamp('   OMS: Linked to id: ' + caseid_nr)

        # Check here if case ID and case code match based on SOC?

        # Check the status
        current_status_oms, current_status_oms_index = get_current_status_line('overview')
        # Check if we need to pass by the OMS or not
        if current_status_oms == destination_status_oms:
            # If the status is already on the correct status, we do not need to go through the OMS
            promote_in_oms = False
        else:
            if current_status_oms_index < destination_status_oms_index:
                # If the current status index is lower destination index, than we need to promote in the OMS
                promote_in_oms = True
            else:
                # If not lower, than do nothing (probably the status is already further)
                promote_in_oms = False

        # Check if it really is a Livit case
        isIndeedLivit = False
        if shouldBeLivit:
            # xpathsearch_companyname = '/html/body/form/div[3]/div[3]/div/div[2]/div[2]/div/div[1]/div[2]/dl[1]/dd[2]/a'
            xpathsearch_companyname = '/html/body/form/div[3]/div[3]/div/div/div[2]/div/div[3]/div[3]/div/table/tbody/tr[2]/td[6]'
            if driver.find_element(By.XPATH, xpathsearch_companyname).text == 'Livit Orthopedie bv Company':
                isIndeedLivit = True
            else:
                isIndeedLivit = False
                # Show error message
                messagebox.showwarning("Warning", caseid + " is not a Livit case. Please follow up that the case is reprinted at MTLS. Inform the real customer about a possible delay.")
                #register_case_ID('invalid',caseid + ' (not Livit)',cc)
                caseids_summary[caseid]['OMS'] = 'Invalid (not Livit)'

        if "Cancelled" not in current_status_oms and promote_in_oms == True: 
            if (shouldBeLivit == True and isIndeedLivit == True) or (shouldBeLivit == False and isIndeedLivit == False):
                # Get the current status of the case ID
                # current_status_oms, current_status_oms_index = get_current_status()
                first_time_batch = 1
                first_production_run = True
                new_loop = True
                while current_status_oms_index < destination_status_oms_index and no_error and new_loop:
                    if statusFlowOms[current_status_oms_index] == 'Production' and first_production_run:
                        first_production_run = False
                        # Grab the element with the just searched id and open a new tab
                        xpathsearch = '/html/body/form/div[3]/div[3]/div/div/div[2]/div/div[3]/div[3]/div/table/tbody/tr[2]/td[2]'
                        driver.find_element(By.XPATH, xpathsearch).click()
                        driver.switch_to.window(driver.window_handles[handle_oms_order_detail])
                        time.sleep(0.1)
                        # Check if a random element exists on the current order page, to check if it is loaded
                        xpath_currentpage = '/html/body/form/div[3]/div[3]/div/div[2]/div[1]/div[1]/span'
                        wait_until_element_is_present('xpath',xpath_currentpage,20)
                        # Check if you need to take the order or not
                        take_order_button_id = 'take-case-button'
                        take_order_button_present = check_exists_by_id(take_order_button_id)
                        if take_order_button_present:
                            print_with_timestamp('   OMS: Clicking the "Take order" button.')
                            button = driver.find_element(By.ID, take_order_button_id)
                            time.sleep(0.1)
                            button.click()
                            time.sleep(0.1)

                        print_with_timestamp('   OMS: The destination status is ' + destination_status_oms + ' (' + str(destination_status_oms_index) + '). We need to further promote this case.')
                        # Click the Parts List tab
                        element_status2 = False
                        xpath_partslist_active = '/html/body/form/div[3]/div[3]/div/div[2]/div[2]/div/ul/li[6]'
                        # xpath_kit = '//*[@id="jqgh_ctl00_right_side_tc_tabPartsList_PartGrid_KitCode"]'
                        while element_status2 == False:
                            # element_status2 = check_exists_by_type('xpath', xpath_kit, 20)
                            xpathsearch_partslist_button = '/html/body/form/div[3]/div[3]/div/div[2]/div[2]/div/ul/li[6]/a'
                            wait_until_element_is_present('xpath',xpathsearch_partslist_button,20)
                            time.sleep(0.1)
                            button = driver.find_element(By.XPATH, xpathsearch_partslist_button)
                            time.sleep(1)
                            button.click()
                            time.sleep(0.1)

                            classes = driver.find_element(By.XPATH, xpath_partslist_active).get_attribute('class')
                            if 'ui-state-active' in classes:
                                element_status2 = True


                        # Check if it is on Streamics(built) or not
                        current_production_substatus = get_production_substatus()
                        if current_production_substatus != 'Streamics (Built)':
                            # Click the update all button
                            xpathsearch_update_allbutton = '/html/body/form/div[3]/div[3]/div/div[2]/div[2]/div/div[6]/div[5]/div/input[6]'
                            wait_until_element_is_present('xpath',xpathsearch_update_allbutton,20)
                            button = driver.find_element(By.XPATH, xpathsearch_update_allbutton)
                            time.sleep(1)
                            button.click()
                            time.sleep(1) # 5
                            print_with_timestamp('   OMS: The case should be promoted to Streamics (Built)')

                        
                        print_with_timestamp('   OMS: The destination status is ' + destination_status_oms + ' (' + str(destination_status_oms_index) + '). We need to further promote this case.')
                        """   # Click the promotion button
                        xpathsearch_promotionbutton = '/html/body/form/div[3]/div[3]/div/div[2]/div[4]/div/div[2]/div[3]/button'
                        element_status = False
                        while element_status == False:
                            element_status = check_exists_by_type('xpath', xpathsearch_promotionbutton, 20)
                        button = driver.find_element(By.XPATH, xpathsearch_promotionbutton)
                        time.sleep(3)
                        button.click()
                        time.sleep(0.1)
                        # Click the confirm button
                        xpathsearch_confirmbutton = '/html/body/form/div[15]/div[2]/div/button[1]'
                        element_status = False
                        while element_status == False:
                            element_status = check_exists_by_type('xpath', xpathsearch_confirmbutton, 20)
                        time.sleep(3)
                        button = driver.find_element(By.XPATH, xpathsearch_confirmbutton)
                        time.sleep(0.1)
                        button.click()
                        time.sleep(1) # 20

                        # Verify if a error message appears
                        error_button_id = 'ui-dialog-title-ctl00_right_side_ucStatusControl_PromotionErrorMessage'
                        xpathsearch_error_button = "/html/body/form/div[16]"
                        xpathsearch_error_button_specific = "/html/body/form/div[16]/div[2]/div/button"
                        error_button_present = driver.find_element(By.XPATH, xpathsearch_error_button).value_of_css_property("display") 
                        
                        if error_button_present != "none":
                            no_error = False
                            print_with_timestamp('   The case could not be promoted.')                
                            print_with_timestamp('   Clicking the error button.')
                            driver.find_element(By.XPATH, xpathsearch_error_button_specific).click()
                            time.sleep(1) # 5
                            register_case_ID('invalid',caseid + ' (invalid production status)',cc)
                            # Close the tab and go back to the overview tab
                        
                        if no_error:
                            current_status_oms, current_status_oms_index = get_current_status() """
                        driver.close()
                        driver.switch_to.window(driver.window_handles[handle_oms_view_orders])
                        time.sleep(0.1)
                        
                    
                    if (current_status_oms_index >= statusFlowOms.index('Production')) and (current_status_oms_index < destination_status_oms_index ):
                        print_with_timestamp('   OMS: The destination status is ' + destination_status_oms + ' (' + str(destination_status_oms_index) + '). We need to further promote this case.')
                        # Go to the batch promotion tab
                        driver.switch_to.window(driver.window_handles[handle_oms_batch_promotion])
                        time.sleep(0.5)
                        # Check if there is still a confirmation button from the last case that failed. If yes, close it first.
                        try:
                            promotion_button_confirm3 = '/html/body/form/div[11]/div[2]/div/button'
                            driver.find_element(By.XPATH, promotion_button_confirm3).click()
                        except:
                            DoNothing = 1

                        time.sleep(1.5)
                        if first_time_batch == 1:
                            wait_until_element_is_present('id','gs_CaseID',20)
                            time.sleep(0.1)
                        # Clear the case ID search field
                        driver.find_element(By.ID, "gs_CaseID").clear()
                        # Grab the case ID search field
                        searchinput = driver.find_element(By.ID, "gs_CaseID")
                        # Insert the case ID into the case ID search field
                        searchinput.send_keys(caseid)
                        time.sleep(1)
                        searchinput.send_keys(Keys.ENTER)
                        
                        # Click the checkbox of the case
                        time.sleep(2)
                        checkbox_ID = "jqg_ctl00_right_side_GridBatch_" + str(caseid_nr)
                        wait_until_element_is_present('id',checkbox_ID,20)
                        time.sleep(1)
                        driver.find_element(By.ID, checkbox_ID).click()
                        time.sleep(1)
                        # Click the promotion button
                        promotion_button_confirm = 'ctl00_right_side_btnPromote'
                        wait_until_element_is_present('id',promotion_button_confirm,20)
                        time.sleep(0.1)
                        driver.find_element(By.ID, promotion_button_confirm).click()
                        time.sleep(1)
                        # Confirm the promotion
                        promotion_button_confirm2 = '/html/body/form/div[8]/div[2]/div/button[1]'
                        wait_until_element_is_present('xpath',promotion_button_confirm2,20)
                        time.sleep(0.1)
                        driver.find_element(By.XPATH, promotion_button_confirm2).click()
                        time.sleep(7)
                        # Confirm the promotion again
                        xpath_conformation_text = '/html/body/form/div[11]/span/div/div'
                        confirmation_text = driver.find_element(By.XPATH, xpath_conformation_text).text
                        if confirmation_text == '0 were promoted successfully, 1 were not promoted':
                            new_loop = False
                            no_error = False

                        promotion_button_confirm3 = '/html/body/form/div[11]/div[2]/div/button'
                        wait_until_element_is_present('xpath',promotion_button_confirm3,20)
                        time.sleep(1)
                        driver.find_element(By.XPATH, promotion_button_confirm3).click()
                        time.sleep(1)

                        # Check if there is still a confirmation button from the last case that failed. If yes, close it first.
                        try:
                            driver.find_element(By.XPATH, promotion_button_confirm3).click()
                        except:
                            DoNothing = 1
                        # 
                        try:
                            driver.find_element(By.ID, checkbox_ID)
                        except:
                            new_loop = False

    
                        

                    # Get the status again
                    if no_error and new_loop:
                        if destination_status_oms == 'Built' and current_status_oms == 'Built':
                            # Do nothing - Cases is in the correct status, we need to proceed to the next case
                            DoNothing = 1
                        else:
                            if current_status_oms_index >= statusFlowOms.index('Production'):
                                current_status_oms, current_status_oms_index = get_current_status_line('batch')
                                first_time_batch = 0
                            else:
                                current_status_oms, current_status_oms_index = get_current_status()
                
                    
                if no_error:
                    print_with_timestamp('   OMS: The destination status ' + destination_status_oms + ' has been reached.')
                    # register_case_ID('valid',caseid,cc)
                    caseids_summary[caseid]['OMS'] = 'Valid'
                else:
                    # register_case_ID('invalid',caseid,cc)
                    caseids_summary[caseid]['OMS'] = 'Invalid'

                
                if cc <= len(caseids)+1:
                    # Go back to the overview of all cases
                    driver.switch_to.window(driver.window_handles[handle_oms_view_orders])
                    time.sleep(0.1)
        else:
            if "Cancelled" in current_status_oms:
                print_with_timestamp('   OMS: This case is in cancelled status (' + current_status_oms + '). Added to invalid cases.')
                #register_case_ID('invalid',caseid + " (cancelled)",cc)
                caseids_summary[caseid]['OMS'] = 'Invalid (cancelled)'
            if promote_in_oms == False and "Cancelled" not in current_status_oms:
                print_with_timestamp('   OMS: This case is already has the correct status (' + current_status_oms + '). Added to valid cases.')
                #register_case_ID('valid',caseid,cc)
                caseids_summary[caseid]['OMS'] = 'Valid'
            time.sleep(0.1)
            driver.switch_to.window(driver.window_handles[handle_oms_view_orders])
        try:
            driver.switch_to.window(driver.window_handles[handle_oms_order_detail])
            driver.close()
        except:
            DoNothing = 1
        driver.switch_to.window(driver.window_handles[handle_oms_view_orders])
        time.sleep(0.1)
        del caseid_nr

    # Add the final time stamp
    timestamps.append(datetime.now)
    # Calculate the average time


    # print_with_timestamp(' ')
    # print_with_timestamp('Valid OMS promotions (' + str(len(oms_caseids_valid)) + '/' + str(len(caseids)) + ')')
    # print_with_timestamp('----------------------------')
    # print_casebycase(oms_caseids_valid)
    # print_with_timestamp(' ')
    # print_with_timestamp('Invalid OMS promotions (' + str(len(oms_caseids_invalid)) + '/' + str(len(caseids)) + ')')
    # print_with_timestamp('----------------------------')
    # print_casebycase(oms_caseids_invalid)
    # print_with_timestamp(' ')

    print_with_timestamp(' ')
    print_with_timestamp(' ')
    print_with_timestamp('---------------------------------------------------------------------- ')
    print_with_timestamp('---------------------------------------------------------------------- ')
    print_with_timestamp('---------------------------------------------------------------------- ')
    print_with_timestamp(' ')
    print_with_timestamp('Full summary')
    print_summary(caseids_summary,caseids_rebuilt_summary,False)
    print_with_timestamp(' ')
    print_with_timestamp('---------------------------------------------------------------------- ')
    print_with_timestamp('---------------------------------------------------------------------- ')
    print_with_timestamp('---------------------------------------------------------------------- ')
    print_with_timestamp(' ')
    print_with_timestamp('Take action for the following cases')
    print_summary(caseids_summary,caseids_rebuilt_summary,True)
    print_with_timestamp(' ')

else:
    print_with_timestamp('No case IDs present.')

print(' ')
print('Script finished.')
print(' ')
# outputfile.close()


# Some code for testing
test_caseid = ('RS22-SAN-DER','RS22-LOD-EKE','RS22-NAT-ASA','RS22-PIT-JAN')
caseids_summary = {}
for test in test_caseid:
    # print(test)
    caseids_summary[test] = {}
    # caseids_summary[test]['Streamics'] = '-'
    # caseids_summary[test]['OMS'] = '-'
caseids_summary['RS22-SAN-DER']['OMS'] = 'Valid'
caseids_summary['RS22-SAN-DER']['Streamics'] = 'Valid'
caseids_summary['RS22-LOD-EKE']['OMS'] = 'Invalid'
caseids_summary['RS22-LOD-EKE']['Streamics'] = 'Valid'
caseids_summary['RS22-NAT-ASA']['OMS'] = 'Valid'
caseids_summary['RS22-NAT-ASA']['Streamics'] = 'Invalid'
caseids_summary['RS22-PIT-JAN']['OMS'] = 'Invalid'
caseids_summary['RS22-PIT-JAN']['Streamics'] = 'Invalid'