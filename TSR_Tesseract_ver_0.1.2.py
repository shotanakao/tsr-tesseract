'''
Python ver 3.6.8
Written by Shota Nakao
'''
'''
importing modules
'''
from tkinter import filedialog
from tkinter import Tk
import os
import pyautogui
import pytesseract
import mss
import mss.tools
import csv
from PIL import Image
import xlsxwriter
import pandas as pd
import time

'''
Asking how many TSRs there are
'''
print('''Hello!
Please type how many TSRs you have and hit the [Enter] key.
(Enter a half-width(半角) number)''')

pageNumber = int(input())
if pageNumber > 0 :
    pageNumber -= 1

'''
Selecting and opening the PDF file
'''
#Open File Select Dialog
root = Tk()
root.filename = filedialog.askopenfilename(initialdir="C:", title="Select file",
                                           filetypes=(("pdf files", "*.pdf"), ("all files", "*.*")))
#Open the PDF
os.startfile(root.filename)

#Using time.sleep to give time for PDF to open before Screenshot
time.sleep(.5)
pyautogui.press('home')

'''
Function to set up SumatraPDF view to zoom in and align regions for screenshotting first section
'''


def pagesetTop():
    pyautogui.keyDown('ctrlleft');
    pyautogui.press('0');
    pyautogui.keyUp('ctrlleft')
    pyautogui.keyDown('ctrlleft');
    pyautogui.press('y');
    pyautogui.keyUp('ctrlleft')
    pyautogui.typewrite('800%')
    pyautogui.press('enter')
    pyautogui.press(['pagedown'])
    pyautogui.press(['right'] * 104)
    pyautogui.press(['down'] * 12)

'''
Function to set up SumatraPDF view to zoom in and align regions for screenshotting second section
'''

def pagesetBottom():
    pyautogui.keyDown('ctrlleft');
    pyautogui.press('0');
    pyautogui.keyUp('ctrlleft')
    pyautogui.keyDown('ctrlleft');
    pyautogui.press('y');
    pyautogui.keyUp('ctrlleft')
    pyautogui.typewrite('230%')
    pyautogui.press('enter')
    pyautogui.press(['down'] * 33)
    pyautogui.press(['right'] * 30)

'''
Function to grab data from each PDF page
'''
def pagedataGrab():   
    '''
    Screenshot Phone and Name portions
    region argument (a 4-integer tuple of (left, top, width, height))
    '''
    pagesetTop()
    time.sleep(.100)
    with mss.mss() as sct:
        # The screen part to capture
        monitor = {"top": 459, "left": 1, "width": 1900, "height": 118}
        output = "sct-{top}x{left}_{width}x{height}.png".format(**monitor)

        # Grab the data
        sct_img = sct.grab(monitor)
        ssName = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")

    with mss.mss() as sct:
        # The screen part to capture
        monitor = {"top": 902, "left": 1, "width": 604, "height": 118}
        output = "sct-{top}x{left}_{width}x{height}.png".format(**monitor)

        # Grab the data
        sct_img = sct.grab(monitor)
        ssPhone = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")
 
        
    pagesetBottom()
    #Screenshotting everything else
    with mss.mss() as sct:
        # The screen part to capture
        monitor = {"top": 316, "left": 1, "width": 804, "height": 108}
        output = "sct-{top}x{left}_{width}x{height}.png".format(**monitor)

        # Grab the data
        sct_img = sct.grab(monitor)
        ssBus = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")
    with mss.mss() as sct:
        # The screen part to capture
        monitor = {"top": 824, "left": 1047, "width": 830, "height": 108}
        output = "sct-{top}x{left}_{width}x{height}.png".format(**monitor)

        # Grab the data
        sct_img = sct.grab(monitor)
        ssOver = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")
    with mss.mss() as sct:
        # The screen part to capture
        monitor = {"top": 572, "left": 1, "width": 804, "height": 108}
        output = "sct-{top}x{left}_{width}x{height}.png".format(**monitor)

        # Grab the data
        sct_img = sct.grab(monitor)
        ssOwner = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")
    with mss.mss() as sct:
        # The screen part to capture
        monitor = {"top": 444, "left": 1047, "width": 830, "height": 108}
        output = "sct-{top}x{left}_{width}x{height}.png".format(**monitor)

        # Grab the data
        sct_img = sct.grab(monitor)
        ssVendor = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")
    with mss.mss() as sct:
        # The screen part to capture
        monitor = {"top": 572, "left": 1047, "width": 830, "height": 108}
        output = "sct-{top}x{left}_{width}x{height}.png".format(**monitor)

        # Grab the data
        sct_img = sct.grab(monitor)
        ssCust = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")

    '''
    No idea why but pytesseract shits itself and can't find tesseract, seems to be in PATH as it should. Manually telling it where it is here.
    '''
    pytesseract.pytesseract.tesseract_cmd = 'c:\\Program Files (x86)\\Tesseract-OCR\\tesseract.exe'
    '''
    OCR w/ tesseract to grab string value
    '''
    txtName = pytesseract.image_to_string(ssName, lang='jpn', config='--psm 7')
    txtPhone = pytesseract.image_to_string(ssPhone, lang='eng', config='--psm 7')
    txtBus = pytesseract.image_to_string(ssBus, lang='jpn', config='--psm 6')
    txtOver = pytesseract.image_to_string(ssOver, lang='jpn', config='--psm 6')
    txtOwner = pytesseract.image_to_string(ssOwner, lang='jpn', config='--psm 6')
    txtVendor = pytesseract.image_to_string(ssVendor, lang='jpn', config='--psm 6')
    txtCust = pytesseract.image_to_string(ssCust, lang='jpn', config='--psm 6')
    
  
    '''
    Storing all of the values in a list
    '''
    pageData = [txtName, txtPhone, txtBus, txtOver,txtOwner,txtVendor,txtCust]

    '''
    Clean up pageData
    '''
    for index, data in enumerate(pageData):
        
        data = data.replace('\n','')
        data = data.replace(u' ','')
        data = data.replace(u'⑨','9')
        data = data.replace(u'⑧','8')
        data = data.replace(u'⑦','7')
        data = data.replace(u'⑥','6')
        data = data.replace(u'⑤','5')
        data = data.replace(u'④','4')
        data = data.replace(u'③','3')
        data = data.replace(u'②','2')
        data = data.replace(u't1on','tion')
        data = data.replace(u'T1ON','TION')
        data = data.replace(u'/','1')
        data = data.replace(u'raもion','ration')
        data = data.replace(u'エ業','工業')
        pageData[index] = data.replace(u'①','1')
    '''
    Adding data from this page to the master list of lists
    '''
    alltsrData.append(pageData)

    '''
    Key pressses to get to the next page
    '''
    pyautogui.keyDown('ctrlleft');pyautogui.press('0');pyautogui.keyUp('ctrlleft')
    pyautogui.press('right')

'''
The list used to store the data OCR'd from the PDFs
'''
alltsrData = [
    ['Company Name', 'Phone #', 'Description', 'Overview', 'Owners', 'Vendors', 'Customers'] 
    ]

'''
Looping through pages to grab data and stick it into the list
'''
while pageNumber > -1:
    pagedataGrab()
    pageNumber -= 1

'''
Saving the list of lists as an excel file
'''
dataFrame = pd.DataFrame.from_records(alltsrData)
dataFrame.to_excel('TSR_excel.xlsx')


