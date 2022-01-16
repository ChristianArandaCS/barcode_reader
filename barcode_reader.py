# this file reads barcodes of all images fed into the script
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
for filename in os.listdir(r'C:\Users\chris\Desktop\python\barcode_reader\imgs'):

    # get img name
    worksheet.write(row, col, filename)

    col = col +1

    # barcode decoder
    img = cv2.imread(filename)
    for code in decode(img):
        detectedBarcodes = code.data.decode('utf-8')

    #i = 0

    # initiate counter for list
    #dataText = detectedBarcodes.data[i]

    # adds information into an excel
    worksheet.write(row, col, detectedBarcodes)

    #i = i + 1

    # row col increments
    row = row +1
    col = col - 1

    
workbook.close()

quit()
