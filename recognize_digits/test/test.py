# -*- coding: utf-8 -*-



import os, sys
import time
import  cv2
from imutils.perspective import four_point_transform
from imutils import contours
import imutils

image = cv2.imread("../example.jpg")
cv2.imshow("test",image)
cv2.waitKey(2000)

image = imutils.resize(image, height=500)
cv2.imshow("test",image)
cv2.waitKey(2000)

gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
cv2.imshow("test",gray)
cv2.waitKey(2000)

blurred = cv2.GaussianBlur(gray, (5, 5), 0)
cv2.imshow("test",blurred)
cv2.waitKey(2000)

edged = cv2.Canny(blurred, 50, 200, 255)
cv2.imshow("test",edged)
cv2.waitKey(2000)

if __name__ == "__main__":
    print(os.path.abspath("__file__"))
    print(os.getcwd())