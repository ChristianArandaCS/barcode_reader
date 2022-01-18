# barcodeRader v1.2
# this file reads barcodes of all images fed into the script
# primarly used to better track and process returns in bulk

# import libraries
import xlsxwriter
import os
import cv2
import shutil
from pathlib import Path
from pyzbar.pyzbar import decode


# set paths
img_path = r'C:\Users\chris\OneDrive\Desktop\scipts\python\barcode_reader\imgs'
script_path = r'C:\Users\chris\OneDrive\Desktop\scipts\python\barcode_reader'


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


# enters for loop to iterate through files
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


# deletes all images in script path directory


quit()
