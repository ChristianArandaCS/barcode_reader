# barcodeRader V2
#  this file reads barcodes of all images fed into the script
# primarly used for returns with barcodes in the images of return shipment label


# import libraries

#from PIL import Image, ImageEnhance, ImageFilter, ImageOps
import xlsxwriter
import os
import cv2
from pyzbar.pyzbar import decode


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


# enters for loop to iterate through files
for filename in os.listdir(r'C:\Users\chris\OneDrive\Desktop\scipts\python\barcode_reader\imgs'):

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

  





    
workbook.close()

quit()
