# barcodeRader v1.4
# 01/18/2022
####################
# this file reads barcodes of all images fed into the script
# primarly used to better track and process returns in bulk

# import libraries
import xlsxwriter
import os
import cv2
import shutil
#from PIL import Image
import PIL.Image
from PIL import ImageDraw
from PIL import ImageFont
from pathlib import Path
from pyzbar.pyzbar import decode
from tkinter import *
from tkinter import filedialog


# set paths
# img_path = r'Z:\Shared\Office Files - Los Angeles Branch\Returns\01.18.22\sebastianImgs\original'
# folder_path = r'Z:\Shared\Office Files - Los Angeles Branch\Returns\01.18.22\sebastianImgs'
# script dir is static
# script_path = r'C:\Users\chris\OneDrive\Desktop\scipts\python\barcode_reader'

script_path = os.path.abspath(os.curdir)

##########################################################################

#creating main root
root = Tk()

root.title("Barcode reader")

##########################################################################

# this function asks the user for the file location
def selectingFromFolder():
    # promts used for location where files are stored
    userPath = filedialog.askdirectory()
    global img_path
    global folder_path
    img_path = userPath
    folder_path = userPath


def runMainScript():

    #sets up excel file
    # generates output excel file
    workbook = xlsxwriter.Workbook('result.xlsx')
    worksheet = workbook.add_worksheet()

    # headers values
    header1 = 'File name'
    header2 = 'Output file'

    # adds headers to the excel
    worksheet.write(0,0,header1)
    worksheet.write(0,1,header2)

    # sets where the script will start from
    row = 1
    col = 0


    # get all file names from img source location
    img_list = os.listdir(img_path)

    # copies a files from the img source to the script path
    for fname in img_list:
        shutil.copy2(os.path.join(img_path,fname),script_path)


    # enters for loop to iterate through files and get barcodes
    for filename in os.listdir(img_path):

        # get img name
        worksheet.write(row, col, filename)

        col = col +1

        # barcode decoder
        img = cv2.imread(filename)
        for code in decode(img):
            detectedBarcodes = code.data.decode('utf-8')

            # adds information into an excel
            worksheet.write(row, col, detectedBarcodes)
            # row col increments

            col = col + 1

            # num of columns to go back
            i = 1
            i = i + 1

        # sets and resets the rows and cols in excel
        row = row +1
        col = 0

    # finished adding barcodes to excel and closes excel    
    workbook.close()


    # create a folder in the image file location
    labeled_imgs = os.path.join(folder_path,"labeled_imgs")
    os.mkdir(labeled_imgs)


    # labels all images in the folder "labeled_imgs"
    for labeled_img in img_list:

        # Open an Image
        img_labeled = PIL.Image.open(labeled_img)
        
        # Call draw Method to add 2D graphics in an image
        I1 = ImageDraw.Draw(img_labeled)
        
        # Custom font style and font size
        myFont = ImageFont.truetype('arial.ttf', 65)
        
        # Add Text to an image
        I1.text((50, 100), labeled_img, font=myFont ,fill=0)
        
        # Save the edited image
        img_labeled.save(labeled_img)


    # move all labeled images from script dir to img dir
    for fname in img_list:
        shutil.move(os.path.join(script_path,fname), labeled_imgs)


    # create ppaths for the excel to move to
    src_excel = os.path.join(script_path,'result.xlsx')
    dest_excel = os.path.join(folder_path,'result.xlsx')

    # move excel to folder dir
    dest = shutil.move(src_excel, dest_excel)


##########################################################################
#creating GUI widgets
scriptDescript = Label(root, text="This script reads all images in a folder \n for barcodes and outputs them in a specified folder")
label_1 = Label(root, text="Folder: ", padx=10, pady=15)
folderSelect = Button(root, text="Select", width=10, borderwidth=3, command=selectingFromFolder)
runScript = Button(root, text="Run script", width=10, borderwidth=3, command=runMainScript)

scriptDescript.grid(row=0, column=0, columnspan=2)
label_1.grid(row=1, column=0)
folderSelect.grid(row=1, column=1)
runScript.grid(row=2, column=1)


##########################################################################
root.mainloop()

quit()

### A DECISION FOR SOMETING IS A DECISION AGAINST SOMETHING ELSE
### 4.669