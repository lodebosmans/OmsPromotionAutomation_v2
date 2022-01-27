#pip install selenium
#pip install update
#pip install chromedriver

# Add the Chromedriver in the same folder als the python script
# https://sites.google.com/a/chromium.org/chromedriver/downloads
# https://stackoverflow.com/questions/40555930/selenium-chromedriver-executable-needs-to-be-in-path


# Tutorial
# https://medium.com/python-in-plain-english/create-your-browser-automation-robot-with-python-and-selenium-ed0db1d6d65d

# email = ""
# password = ""

import time 
import os
import re
import sys
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from tkinter import *
from tkinter.ttk import *
from tkinter import messagebox
from datetime import datetime

def click_on_view_orders():
    print_with_timestamp('--------------------------------------------------------------------------------------------')
    print_with_timestamp(' ')
    print_with_timestamp('Focus on the "view orders" button and click it')
    print_with_timestamp(' ')
    xpathsearch = '//*[@id="ctl00_left_side_ucMm_accM_nwMm"]/ul/li[2]/a'
    button = driver.find_element(By.XPATH, xpathsearch)
    button.click()
    element_status = False
    while element_status == False:
        element_status = check_exists_by_type('xpath', '/html/body/form/div[3]/div[3]/div/h1', 20*delay_factor)
    # time.sleep(10*delay_factor)

def click_all_buttons_overview(xpathsearch_firstall,xpathsearch_secondall):
    time.sleep(1*delay_factor)
    print_with_timestamp('Focus on the first "all" button and click it')
    driver.find_element(By.XPATH, xpathsearch_secondall).click()
    print_with_timestamp('Focus on the second "all" button and click it')
    driver.find_element(By.XPATH, xpathsearch_firstall).click()
    print_with_timestamp(' ')
    time.sleep(15*delay_factor)
    if driver.find_element(By.ID, 'ctl00_right_side_openInNewWindow').is_selected() == False:
        driver.find_element(By.ID, "ctl00_right_side_openInNewWindow").click()
        time.sleep(0.5*delay_factor)


def check_exists_by_id(id):
    try:
        driver.find_element(By.ID, id)
        print_with_timestamp('   The element with id ' + id + ' exists. So we need to click the "Take order" button.')
    except:
        print_with_timestamp('   The element with id ' + id + ' does not exist. So we do not need to click the "Take order" button.')
        return False
    return True

def check_exists_by_type(elementtype, element_id_or_path, ticker):
    result = False
    x = 0
    while result == False and x <= ticker:
        x += 1
        if x > ticker:
            print_with_timestamp('   The script seems to be stuck. Consider restarting the script.')
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
        print('None')
    else:
        for case in listcases:
            print(case)

def get_current_status():
    time.sleep(1*delay_factor)
    xpathsearch_statusbar = '/html/body/form/div[3]/div[3]/div/div[2]/div[1]/div[4]/div/span'
    current_status = driver.find_element(By.XPATH, xpathsearch_statusbar).text
    if "Cancelled" not in current_status: 
        current_status_index = statusFlow.index(current_status)
        print_with_timestamp('   The current status is: ' + current_status + ' (' + str(current_status_index) + ')')
    else:
        current_status_index = 'X'
        print_with_timestamp('   The current status is: ' + current_status )
    return current_status, current_status_index

def get_current_status_line(page):
    time.sleep(1*delay_factor)
    if page == 'overview':
        xpathsearch_statusline = '/html/body/form/div[3]/div[3]/div/div/div[2]/div/div[3]/div[3]/div/table/tbody/tr[2]/td[7]/div/div/div[1]/span'
    if page == 'batch':
        xpathsearch_statusline = '/html/body/form/div[3]/div[3]/div/div/div[2]/div[2]/div[3]/div[3]/div/table/tbody/tr[2]/td[9]/div/div/div[1]/span'
    current_status = driver.find_element(By.XPATH, xpathsearch_statusline).text
    if "Cancelled" not in current_status: 
        current_status_index = statusFlow.index(current_status)
        print_with_timestamp('   The current status is: ' + current_status + ' (' + str(current_status_index) + ')')
    else:
        current_status_index = 'X'
        print_with_timestamp('   The current status is: ' + current_status )
    return current_status, current_status_index

def get_production_substatus():
    #time.sleep(10*delay_factor)
    xpathsearch_streamics_status = '/html/body/form/div[3]/div[3]/div/div[2]/div[2]/div/div[6]/div[4]/div[3]/div[3]/div/table/tbody/tr[2]/td[6]'
    element_status = False
    while element_status == False:
        element_status = check_exists_by_type('xpath',xpathsearch_streamics_status,20*delay_factor)
    current_status = driver.find_element(By.XPATH, xpathsearch_streamics_status).text
    print_with_timestamp('   The current production substatus is: ' + current_status )
    return current_status

def register_case_ID(validation_status,caseid,cc):
    if validation_status == 'valid':
        print_with_timestamp('   Registering the valid case promotion.')
        caseids_valid.append(caseid)
    elif validation_status == 'invalid':
        print_with_timestamp('   Registering the invalid case promotion.')
        caseids_invalid.append(caseid)
    else:
        print_with_timestamp(' ')
        print_with_timestamp('Incorrect use of validation function.')
    
    print_with_timestamp(' ')
    print_with_timestamp('These are currently the valid cases:')
    if len(caseids_valid) == 0:
        print_with_timestamp('None')
    else:
        print_with_timestamp(caseids_valid)
    print_with_timestamp(' ')
    print_with_timestamp('These are currently the invalid cases:')
    if len(caseids_invalid) == 0:
        print_with_timestamp('None')
    else:
        print_with_timestamp(caseids_invalid)
    print_with_timestamp(' ')
    print_with_timestamp('These are the remaining cases still to be done:')
    print_with_timestamp(caseids[cc:])
    print_with_timestamp(' ')


def print_with_timestamp(input):
    f=open(logfile, "a+")
    if str(type(input)) != "<class 'list'>":
        # Normal print
        toprint = str(datetime.now()) + '    ' + input
        print(toprint)
        # Print for logfile
        f.write(toprint + '\n')
    else:
        # Normal print
        print(str(datetime.now()) + '    ', end = '')
        print(*input, sep = ", ")  
        # Print for logfile
        listcases = ''
        for case in input:
            listcases = listcases + case + ', '
        listcases = listcases[0:len(listcases)-2]
        f.write(str(datetime.now()) + '    ' + listcases + '\n')

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
logfile = 'log/' + datepieces['y'] + datepieces['mo'] + datepieces['d'] + '_' + datepieces['h'] + datepieces['mi'] + datepieces['s'] + '_logfile.txt'

# Define the status order
statusFlow = ["Waiting for design parameters","Design","Design rejected","Design QC","Production","Built","Ready to ship","Shipped"]

# Define the delay factor
delayFactor = ["1","1.25","1.5","1.75","2","2.25","2.5","2.75","3"]

# Define the width of the input fields
width_inputfield = 40

# Generate the input screen 
def clicked():
    global raw_input_caseids
    global destination_status
    global destination_status_index
    global delay_factor
    global email
    global password
    raw_input_caseids = txt.get("1.0","end") # https://www.delftstack.com/howto/python-tkinter/how-to-get-the-input-from-tkinter-text-box/
    destination_status = combo.get()
    destination_status_index = statusFlow.index(destination_status)
    # delay_factor = float(combo_slower.get())
    delay_factor = 1
    # email = entry_email.get()
    email = entry_email.get()
    password = entry_password.get()
    window.destroy() # Closes the internal loop and lets the script run forward, otherwise it will freeze here.

window = Tk()
window.title("Status promotions")
window.geometry('400x500')


label_spacer0 = Label(window, text="          ")
label_spacer0.grid(column=0, row=0)

label_spacer2 = Label(window, text=" ")
label_spacer2.grid(column=1, row=0)

label_input = Label(window, text="Insert the case IDs:")
label_input.grid(column=1, row=1)
txt = Text(window,width=width_inputfield,height=10)
txt.grid(column=1, row=2)

label_spacer3 = Label(window, text=" ")
label_spacer3.grid(column=1, row=3)

label_combo = Label(window,text="Choose a status:")
label_combo.grid(column=1, row=4)
combo = Combobox(window,width=width_inputfield)
combo['values']= ("Choose a status", "Built", "Ready to ship", "Shipped")
combo.current(0) #set the selected item
combo.grid(column=1, row=5)

label_spacer = Label(window, text=" ")
label_spacer.grid(column=1, row=6)

Livit = IntVar()
Checkbutton(window, text="Livit cases?", variable=Livit).grid(column=1, row=7)

# label_spacer3 = Label(window, text=" ")
# label_spacer3.grid(column=1, row=8)

# label_combo_slower = Label(window,text="Increase waiting time:")
# label_combo_slower.grid(column=1, row=9)
# combo_slower = Combobox(window,width=40)
# combo_slower['values']= delayFactor
# combo_slower.current(0) #set the selected item
# combo_slower.grid(column=1, row=10)

label_spacer4 = Label(window, text=" ")
label_spacer4.grid(column=1, row=11)

label_email = Label(window, text="Email address")
label_email.grid(column=1, row=12)

# entry_email = Entry(window, text="")
# entry_email.grid(column=1, row=13)

entry_email = Combobox(window,width=width_inputfield)
entry_email['values']= ("Choose an email address", "rkia.elhassani@rsprint.be","julie.wellens@materialise.be","laura.janssens@materialise.be" ,"mariska.swolfs@rsprint.be" , "sander.van.nieuwenhoven@rsprint.be", "pieter-jan.lijnen@rsprint.be", "lode.bosmans@rsprint.be","flowbuiltproduction@gmail.com")
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
window.mainloop()
# After the click on the button, the window is destroyed, so data can not be collected again. Check function 'clicked'.


# See if it is for Livit or not
if Livit.get() == 1:
    shouldBeLivit = True
else:
    shouldBeLivit = False

# Get the indices where the case IDs are located
indices = [m.start() for m in re.finditer('RS2', raw_input_caseids)]
caseids = []
# For every case, insert it in the case ID array
for x in indices:
    newcase = raw_input_caseids[x:x+12]
    if newcase not in caseids:
        caseids.append(newcase)

print_with_timestamp('The cases to be processed:')
print_with_timestamp(caseids)
print_with_timestamp(' ')

if len(caseids) > 0:
    # Startup the browser
    chrome_options = webdriver.ChromeOptions()
    driver = webdriver.Chrome('chromedriver',options=chrome_options)

    # Get the website you want to open
    driver.get("https://portal.rsprint.com/Public/Default.aspx")
    # Apply the delay factor twice to the regular timeout, so only the most critical steps are enforced squared
    driver.set_page_load_timeout(500*delay_factor*delay_factor)
    print_with_timestamp(' ')
    print_with_timestamp('Opening the website')

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
    if email != "" and password != "":
        # Press the button
        button = driver.find_element(By.ID, "ctl00_left_side_ucL_loginView_accP_vL_cL_LoginButton")
        button.click()
    if email != "" or password != "":
        # Give some time to insert the credentials and press enter
        time.sleep(15*delay_factor)

    xpath_vieworders = '/html/body/form/div[3]/div[2]/div[1]/div/div[1]/ul/li[2]/a'
    element_status = False
    while element_status == False:
        element_status = check_exists_by_type('xpath',xpath_vieworders,20*delay_factor)
    
    # Get the OMS role
    xpathsearch_OmsRole = 'ctl00_left_side_ucL_loginView_accL_vLi_lblR'
    element_status = False
    while element_status == False:
        element_status = check_exists_by_type('id',xpathsearch_OmsRole,20*delay_factor)
    current_role = driver.find_element(By.ID, xpathsearch_OmsRole).text
    if current_role == 'Mat Admin':
        xpathsearch_secondall = '//*[@id="ctl00_right_side_RadioButtonOwnerFilter"]/label[2]/span'
        xpathsearch_secondall_pressed = '//*[@id="ctl00_right_side_RadioButtonOwnerFilter"]/label[2]'
    else:
        xpathsearch_secondall = '//*[@id="ctl00_right_side_RadioButtonOwnerFilter"]/label[3]/span'
        xpathsearch_secondall_pressed = '//*[@id="ctl00_right_side_RadioButtonOwnerFilter"]/label[3]'
            

    # Focus on the 'view orders' button and click it
    click_on_view_orders()
    time.sleep(8*delay_factor)

    # Focus on the first 'all' button and click it
    # print_with_timestamp('Focus on the first "all" button and click it')
    xpathsearch_firstall = '//*[@id="ctl00_right_side_RadioButtonStatusFilter"]/label[4]/span'
    xpathsearch_firstall_pressed = '//*[@id="ctl00_right_side_RadioButtonStatusFilter"]/label[4]'
    element_status = False
    while element_status == False:
        element_status = check_exists_by_type('xpath',xpathsearch_firstall,20*delay_factor)
    time.sleep(5*delay_factor)
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
    driver.switch_to.window(driver.window_handles[1])
    time.sleep(0.1)
    driver.get('https://portal.rsprint.com/Public/CaseManagement/ViewBatchCaseList.aspx')
    #time.sleep(15*delay_factor)

    # Click the batch promotion all button
    print_with_timestamp('Focus on the batch promotion "all" button and click it')
    print_with_timestamp(' ')
    xpathsearch_batchall = '/html/body/form/div[3]/div[3]/div/div/div[1]/div/span/label[2]/span'
    element_status = False
    while element_status == False:
        element_status = check_exists_by_type('xpath',xpathsearch_batchall,20*delay_factor)
    time.sleep(1*delay_factor)
    driver.find_element(By.XPATH, xpathsearch_batchall).click()
    time.sleep(3*delay_factor)

    # Go back to the view orders screen
    driver.switch_to.window(driver.window_handles[0])
    time.sleep(0.1)


    # Go over all case IDs
    cc = 0
    caseids_valid = []
    caseids_invalid = []
    timestamps = []
    for caseid in caseids:
        # Register the timestamp and other vars
        timestamps.append(datetime.now)
        no_error = True
        cc += 1
        #try:
        # See if the all buttons are still correct
        check_first_all = driver.find_element(By.XPATH, xpathsearch_firstall_pressed).get_attribute('aria-pressed')
        check_second_all = driver.find_element(By.XPATH, xpathsearch_secondall_pressed).get_attribute('aria-pressed')
        if check_first_all != 'true' or check_second_all != 'true':
            click_all_buttons_overview(xpathsearch_firstall,xpathsearch_secondall)
         
        print_with_timestamp('Processing: ' + caseid + ' (' + str(cc) + '/' + str(len(caseids)) + ')')
        # Clear the case ID search field
        element_status = False
        while element_status == False:
            element_status = check_exists_by_type('id','gs_CaseID',20*delay_factor)
        driver.find_element(By.ID, "gs_CaseID").clear()
        # Grab the case ID search field
        searchinput = driver.find_element(By.ID, "gs_CaseID")
        # Insert the case ID into the case ID search field
        searchinput.send_keys(caseid)
        time.sleep(1*delay_factor)
        searchinput.send_keys(Keys.ENTER)
        # Define the Xpath to the case ID in the result table
        xpathsearch_id = '/html/body/form/div[3]/div[3]/div/div/div[2]/div/div[3]/div[3]/div/table/tbody/tr[2]'
        xpathsearch_caseid = '/html/body/form/div[3]/div[3]/div/div/div[2]/div/div[3]/div[3]/div/table/tbody/tr[2]/td[2]'
        element_status = False
        while element_status == False:
            element_status = check_exists_by_value('xpath', xpathsearch_caseid, 30*delay_factor, caseid)
        # Grab the case ID in the result table row
        caseid_nr = driver.find_element(By.XPATH, xpathsearch_id).get_attribute("id")
        print_with_timestamp('   Linked to id: ' + caseid_nr)

        # Check here if case ID and case code match based on SOC?

        # Check the status
        current_status, current_status_index = get_current_status_line('overview')

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
                register_case_ID('invalid',caseid + ' (not Livit)',cc)

        if "Cancelled" not in current_status: 
            if (shouldBeLivit == True and isIndeedLivit == True) or (shouldBeLivit == False and isIndeedLivit == False):
                # Get the current status of the case ID
                # current_status, current_status_index = get_current_status()
                first_time_batch = 1
                first_production_run = True
                new_loop = True
                while current_status_index < destination_status_index and no_error and new_loop:
                    if statusFlow[current_status_index] == 'Production' and first_production_run:
                        first_production_run = False
                        # Grab the element with the just searched id and open a new tab
                        xpathsearch = '/html/body/form/div[3]/div[3]/div/div/div[2]/div/div[3]/div[3]/div/table/tbody/tr[2]/td[2]'
                        driver.find_element(By.XPATH, xpathsearch).click()
                        driver.switch_to.window(driver.window_handles[2])
                        time.sleep(0.1)
                        # Check if a random element exists on the current order page, to check if it is loaded
                        xpath_currentpage = '/html/body/form/div[3]/div[3]/div/div[2]/div[1]/div[1]/span'
                        element_status = False
                        while element_status == False:
                            element_status = check_exists_by_type('xpath', xpath_currentpage, 20*delay_factor)
                        # Check if you need to take the order or not
                        take_order_button_id = 'take-case-button'
                        take_order_button_present = check_exists_by_id(take_order_button_id)
                        if take_order_button_present:
                            print_with_timestamp('   Clicking the "Take order" button.')
                            button = driver.find_element(By.ID, take_order_button_id)
                            time.sleep(0.1*delay_factor)
                            button.click()
                            time.sleep(0.1)

                        print_with_timestamp('   The destination status is ' + destination_status + ' (' + str(destination_status_index) + '). We need to further promote this case.')
                        # Click the Parts List tab
                        element_status2 = False
                        xpath_partslist_active = '/html/body/form/div[3]/div[3]/div/div[2]/div[2]/div/ul/li[6]'
                        # xpath_kit = '//*[@id="jqgh_ctl00_right_side_tc_tabPartsList_PartGrid_KitCode"]'
                        while element_status2 == False:
                            # element_status2 = check_exists_by_type('xpath', xpath_kit, 20*delay_factor)
                            xpathsearch_partslist_button = '/html/body/form/div[3]/div[3]/div/div[2]/div[2]/div/ul/li[6]/a'
                            element_status = False
                            while element_status == False:
                                element_status = check_exists_by_type('xpath', xpathsearch_partslist_button, 20*delay_factor)
                            time.sleep(0.1)
                            button = driver.find_element(By.XPATH, xpathsearch_partslist_button)
                            time.sleep(1*delay_factor)
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
                            element_status = False
                            while element_status == False:
                                element_status = check_exists_by_type('xpath', xpathsearch_update_allbutton, 20*delay_factor)
                            button = driver.find_element(By.XPATH, xpathsearch_update_allbutton)
                            time.sleep(1*delay_factor)
                            button.click()
                            time.sleep(1*delay_factor) # 5
                            print_with_timestamp('   The case should be promoted to Streamics (Built)')

                        
                        print_with_timestamp('   The destination status is ' + destination_status + ' (' + str(destination_status_index) + '). We need to further promote this case.')
                        """   # Click the promotion button
                        xpathsearch_promotionbutton = '/html/body/form/div[3]/div[3]/div/div[2]/div[4]/div/div[2]/div[3]/button'
                        element_status = False
                        while element_status == False:
                            element_status = check_exists_by_type('xpath', xpathsearch_promotionbutton, 20*delay_factor)
                        button = driver.find_element(By.XPATH, xpathsearch_promotionbutton)
                        time.sleep(3*delay_factor)
                        button.click()
                        time.sleep(0.1*delay_factor)
                        # Click the confirm button
                        xpathsearch_confirmbutton = '/html/body/form/div[15]/div[2]/div/button[1]'
                        element_status = False
                        while element_status == False:
                            element_status = check_exists_by_type('xpath', xpathsearch_confirmbutton, 20*delay_factor)
                        time.sleep(3*delay_factor)
                        button = driver.find_element(By.XPATH, xpathsearch_confirmbutton)
                        time.sleep(0.1*delay_factor)
                        button.click()
                        time.sleep(1*delay_factor) # 20

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
                            time.sleep(1*delay_factor) # 5
                            register_case_ID('invalid',caseid + ' (invalid production status)',cc)
                            # Close the tab and go back to the overview tab
                        
                        if no_error:
                            current_status, current_status_index = get_current_status() """
                        driver.close()
                        driver.switch_to.window(driver.window_handles[0])
                        time.sleep(0.1)
                        
                    
                    if (current_status_index >= statusFlow.index('Production')) and (current_status_index < destination_status_index ):
                        print_with_timestamp('   The destination status is ' + destination_status + ' (' + str(destination_status_index) + '). We need to further promote this case.')
                        # Go to the batch promotion tab
                        driver.switch_to.window(driver.window_handles[1])
                        time.sleep(0.5)
                        # Check if there is still a confirmation button from the last case that failed. If yes, close it first.
                        try:
                            promotion_button_confirm3 = '/html/body/form/div[11]/div[2]/div/button'
                            driver.find_element(By.XPATH, promotion_button_confirm3).click()
                        except:
                            DoNothing = 1

                        time.sleep(0.5)
                        if first_time_batch == 1:
                            element_status = False
                            while element_status == False:
                                element_status = check_exists_by_type('id', 'gs_CaseID', 20*delay_factor)
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
                        time.sleep(1*delay_factor)
                        checkbox_ID = "jqg_ctl00_right_side_GridBatch_" + str(caseid_nr)
                        element_status = False
                        while element_status == False:
                            element_status = check_exists_by_type('id', checkbox_ID, 20*delay_factor)
                        time.sleep(2*delay_factor)
                        driver.find_element(By.ID, checkbox_ID).click()
                        time.sleep(1)
                        # Click the promotion button
                        promotion_button_confirm = 'ctl00_right_side_btnPromote'
                        element_status = False
                        while element_status == False:
                            element_status = check_exists_by_type('id', promotion_button_confirm, 20*delay_factor)
                        time.sleep(0.1)
                        driver.find_element(By.ID, promotion_button_confirm).click()
                        time.sleep(1*delay_factor)
                        # Confirm the promotion
                        promotion_button_confirm2 = '/html/body/form/div[8]/div[2]/div/button[1]'
                        element_status = False
                        while element_status == False:
                            element_status = check_exists_by_type('xpath', promotion_button_confirm2, 20*delay_factor)
                        time.sleep(0.1)
                        driver.find_element(By.XPATH, promotion_button_confirm2).click()
                        time.sleep(5*delay_factor)
                        # Confirm the promotion again
                        xpath_conformation_text = '/html/body/form/div[11]/span/div/div'
                        confirmation_text = driver.find_element(By.XPATH, xpath_conformation_text).text
                        if confirmation_text == '0 were promoted successfully, 1 were not promoted':
                            new_loop = False
                            no_error = False

                        promotion_button_confirm3 = '/html/body/form/div[11]/div[2]/div/button'
                        element_status = False
                        while element_status == False:
                            element_status = check_exists_by_type('xpath', promotion_button_confirm3, 20*delay_factor)
                        time.sleep(2*delay_factor)
                        driver.find_element(By.XPATH, promotion_button_confirm3).click()
                        time.sleep(3*delay_factor)

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
                        if destination_status == 'Built' and current_status == 'Built':
                            # Do nothing - Cases is in the correct status, we need to proceed to the next case
                            DoNothing = 1
                        else:
                            if current_status_index >= statusFlow.index('Production'):
                                current_status, current_status_index = get_current_status_line('batch')
                                first_time_batch = 0
                            else:
                                current_status, current_status_index = get_current_status()
                
                    
                if no_error:
                    print_with_timestamp('   The destination status ' + destination_status + ' has been reached.')
                    register_case_ID('valid',caseid,cc)
                else:
                    register_case_ID('invalid',caseid,cc)

                
                if cc <= len(caseids)+1:
                    # Go back to the overview of all cases
                    driver.switch_to.window(driver.window_handles[0])
                    time.sleep(0.1)
        else:
            print_with_timestamp('   This case is in cancelled status (' + current_status + '). Added to invalid cases.')
            register_case_ID('invalid',caseid + " (cancelled)",cc)
            driver.switch_to.window(driver.window_handles[0])
            time.sleep(0.1)
        #except:
            #print_with_timestamp('   This case (' + caseid + ') was timed out. Proceding to the next one.')
            #register_case_ID('invalid',caseid + " (timedout)",cc)
            #try:
            #    driver.switch_to.window(driver.window_handles[2])
            #    driver.close()
            #except:
            #    DoNothing = 1
            #driver.switch_to.window(driver.window_handles[0])
            #time.sleep(0.1)
        try:
            driver.switch_to.window(driver.window_handles[2])
            driver.close()
        except:
            DoNothing = 1
        driver.switch_to.window(driver.window_handles[0])
        time.sleep(0.1)

    # Add the final time stamp
    timestamps.append(datetime.now)
    # Calculate the average time


    print(' ')
    print('Valid promotions (' + str(len(caseids_valid)) + '/' + str(len(caseids)) + ')')
    print('----------------------------')
    print_casebycase(caseids_valid)
    print(' ')
    print('Invalid promotions (' + str(len(caseids_invalid)) + '/' + str(len(caseids)) + ')')
    print('----------------------------')
    print_casebycase(caseids_invalid)
    print(' ')
else:
    print_with_timestamp('No case IDs present.')

print(' ')
print('Script finished.')
print(' ')
# outputfile.close()