# this file reads barcodes of all images fed into the script
# primarly used for returns with barcodes in the images of return shipment label


# import libraries

from PIL import Image, ImageEnhance, ImageFilter, ImageOps
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


# creates the function all images are going to cylce through
def BarcodeReader(filename):

    # read the image in numpy array
    img = cv2.imread(filename)

    # decode barcode in image
    detectedBarcodes = decode(img)

    # if not detected add placeholder into textfile
    if not detectedBarcodes:
        detectedBarcodes = "Barcode not detected my boi"
    else:
        
        # Traverse through all the detected barcodes in image
        for barcode in detectedBarcodes:
        
            # Locate the barcode position in image
            (x, y, w, h) = barcode.rect
            
            # Put the rectangle in image using
            # cv2 to heighlight the barcode
            cv2.rectangle(img, (x-10, y-10),
                        (x + w+10, y + h+10),
                        (255, 0, 0), 2)
            
            if barcode.data!="":
            
            # Print the barcode data
                detectedBarcodes = barcode.data

                return detectedBarcodes





# enters for loop to iterate through files
for filename in os.listdir(r'C:\Users\chris\OneDrive\Desktop\scipts\python\barcode_reader\imgs'):

    # get img name
    worksheet.write(row, col, filename)

    col = col +1

    # barcode decoder


    detectedBarcodes = BarcodeReader(filename)

    # adds information into excel

    # adds information into an excel
    worksheet.write(row, col, detectedBarcodes)

    # row col increments
    row = row +1
    col = col - 1

    
workbook.close()

quit()
