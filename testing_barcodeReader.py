#testing barcode reader

import cv2
from pyzbar.pyzbar import decode

img = cv2.imread('img2.jpeg')
print(decode(img))