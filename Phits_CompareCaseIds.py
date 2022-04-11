import os
import re
from tkinter import *
from tkinter.ttk import *
from datetime import datetime
import win32com.client
import pandas as pd

# -----------------------------------------------------------------------------------------------------------

streamicsOrderFile_path = os.getcwd() + '/input/CaseIdOrderIdMatch.xlsm'

def print_with_timestamp(input):
    f=open(logfile, "a+")
    # Terminal print
    print(input)
    # File print
    f.write(input + '\n')


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


# Generate the input screen 
def clicked():
    global raw_input_caseids_shipmentlist
    global raw_input_caseids_scanned
    global user
    raw_input_caseids_shipmentlist  = txt.get("1.0","end") # https://www.delftstack.com/howto/python-tkinter/how-to-get-the-input-from-tkinter-text-box/
    raw_input_caseids_scanned  = txt_rebuilt.get("1.0","end") # https://www.delftstack.com/howto/python-tkinter/how-to-get-the-input-from-tkinter-text-box/
    window.destroy() # Closes the internal loop and lets the script run forward, otherwise it will freeze here.


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

label_input = Label(window, text="Insert shipment list content")
label_input.grid(column=1, row=1)
txt = Text(window,width=width_inputfield,height=10)
txt.grid(column=1, row=2)

label_spacer5 = Label(window, text=" ")
label_spacer5.grid(column=1, row=14)

btn = Button(window, text="Compare", command=clicked)
btn.grid(column=1, row=18)

# Col 2 and 3

label_spacer0 = Label(window, text="       ")
label_spacer0.grid(column=2, row=1)

label_input_rebuilt = Label(window, text="Insert the scanned case IDs")
label_input_rebuilt.grid(column=3, row=1)
txt_rebuilt = Text(window,width=width_inputfield,height=10)
txt_rebuilt.grid(column=3, row=2)

window.mainloop()
# After the click on the button, the window is destroyed, so data can not be collected again. Check function 'clicked'.

logfile = 'log/' + datepieces['y'] + datepieces['mo'] + datepieces['d'] + '_' + datepieces['h'] + datepieces['mi'] + datepieces['s'] + '_' + '_logfile_comparison_cases.txt'


# Check if the Excel file with the link between case ID and Streamics order ID is present.
if os.path.exists(streamicsOrderFile_path):
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





# Get the cases for rebuilt and for promotion
caseids_shipmentlist = get_caseids_from_input(raw_input_caseids_shipmentlist)
caseids_scanned = get_caseids_from_input(raw_input_caseids_scanned)
# Create the result vars
caseids_in_both = ()
caseids_in_shipmentlist_but_not_in_box = ()
caseids_in_box_bot_not_in_shipmentlist = ()
caseids_that_do_not_exist = ()
# Do the comparison
def compare_cases(origin_cases, target_cases, in_both_file, exception_file, dont_exist):
    for case in origin_cases:
        if case in target_cases:
            if case not in in_both_file:
                in_both_file = in_both_file + (case,)
        else:
            if case not in exception_file:
                exception_file = exception_file + (case,)
        if case not in streamics_order_ids:
            dont_exist = dont_exist + (case,)
    return origin_cases, target_cases, in_both_file, exception_file, dont_exist

caseids_shipmentlist, caseids_scanned, caseids_in_both, caseids_in_shipmentlist_but_not_in_box, caseids_that_do_not_exist = compare_cases(caseids_shipmentlist, caseids_scanned, caseids_in_both, caseids_in_shipmentlist_but_not_in_box, caseids_that_do_not_exist)
caseids_scanned, caseids_shipmentlist, caseids_in_both, caseids_in_box_bot_not_in_shipmentlist, caseids_that_do_not_exist = compare_cases(caseids_scanned, caseids_shipmentlist, caseids_in_both, caseids_in_box_bot_not_in_shipmentlist, caseids_that_do_not_exist)

print_with_timestamp(' ')
print_with_timestamp('Cases in the shipmentlist')
print_with_timestamp('------------------------')
for case in caseids_shipmentlist:
    print_with_timestamp(case)
print_with_timestamp(' ')
print_with_timestamp('Cases in the box')
print_with_timestamp('---------------')
for case in caseids_scanned:
    print_with_timestamp(case)
print_with_timestamp(' ')
print_with_timestamp('Cases that do not exist')
print_with_timestamp('----------------------')
for case in caseids_that_do_not_exist:
    print_with_timestamp(case)
print_with_timestamp(' ')
print_with_timestamp('Cases in the both')
print_with_timestamp('----------------')
for case in caseids_in_both:
    print_with_timestamp(case)
print_with_timestamp(' ')
print_with_timestamp('Cases in the shipmentlist, but not in the box')
print_with_timestamp('--------------------------------------------')
for case in caseids_in_shipmentlist_but_not_in_box:
    print_with_timestamp(case)
print_with_timestamp(' ')
print_with_timestamp('Cases in the box, but not in the shipmentlist')
print_with_timestamp('--------------------------------------------')
for case in caseids_in_box_bot_not_in_shipmentlist:
    print_with_timestamp(case)
print_with_timestamp(' ')

print(' ')
print('Script finished.')
print(' ')

# input("prompt: ")
