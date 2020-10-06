import cv2
import pytesseract
import numpy as np
from datetime import datetime
import pandas as pd
from pandas import ExcelWriter

# get grayscale image
def get_grayscale(image):
    return cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

# noise removal
def remove_noise(image):
    return cv2.medianBlur(image,5)

#thresholding
def thresholding(image):
    return cv2.threshold(image, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]

#dilation
def dilate(image):
    kernel = np.ones((5,5),np.uint8)
    return cv2.dilate(image, kernel, iterations = 1)

#erosion
def erode(image):
    kernel = np.ones((5,5),np.uint8)
    return cv2.erode(image, kernel, iterations = 1)

#opening - erosion followed by dilation
def opening(image):
    kernel = np.ones((5,5),np.uint8)
    return cv2.morphologyEx(image, cv2.MORPH_OPEN, kernel)

#canny edge detection
def canny(image):
    return cv2.Canny(image, 100, 200)

#skew correction
def deskew(image):
    coords = np.column_stack(np.where(image > 0))
    angle = cv2.minAreaRect(coords)[-1]
    if angle < -45:
        angle = -(90 + angle)
    else:
        angle = -angle
    (h, w) = image.shape[:2]
    center = (w // 2, h // 2)
    M = cv2.getRotationMatrix2D(center, angle, 1.0)
    rotated = cv2.warpAffine(image, M, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)
    return rotated

#template matching
def match_template(image, template):
    return cv2.matchTemplate(image, template, cv2.TM_CCOEFF_NORMED)

def dateFormat():
    dater = str(datetime.date(datetime.now()))
    dater = dater.replace("-","/")
    return dater
def checkAssistance(name,list):
    for i in list:
        if i==name:
            print("found " + name)
            return True
    return False
def row_style(row):
    print("Row index: " + str(pos.index))
    if pos.index%2==0:
        return pd.Series('background-color: gray', row.index)


date = dateFormat()
print(date)
img = cv2.imread('imgTest3.png')


#gray = get_grayscale(img)
#thresh = thresholding(gray)
#opening = opening(gray)
#canny = canny(gray)

# Adding custom options
custom_config = r'--oem 3 --psm 6'
s = pytesseract.image_to_string(img, config=custom_config)
arrNames = s.splitlines()
participants = []
for i in arrNames:
    if i!='':

        temp = i.split(' ')
        #print(i)
        #print(temp[1])
        participants.append(temp[1] + ' ' + temp[2])
    
df = pd.ExcelFile('asistencias.xlsx')

dfSheetNames = df.sheet_names
dfLists =[]
n=0
for i in dfSheetNames:
    dfNow = pd.read_excel('asistencias.xlsx',sheet_name=i)
    dfLists.append(dfNow)

    n+=1
    #print(n)
listOHR = []
listName = []
listStatus =[]
listDate = []
listComment = []
for index, row in dfLists[1].iterrows():
    listOHR.append(row['OHR ID'])
    listName.append(row['Name'])
    if checkAssistance(row['Name'],participants):
        listStatus.append("ok")
    else:
        listStatus.append("-")
    listDate.append(date)
    listComment.append('')

testy = pd.DataFrame({"OHR ID":listOHR,"Name":listName,"Date": listDate,"Status":listStatus,"Comments":listComment})
dfLists[0] = dfLists[0].append(testy,ignore_index=True)

n=-1
for index, row in dfLists[0].iterrows():
    n+=1
    dfLists[0].style.apply(row_style, axis=1)



with ExcelWriter('asistencias.xlsx') as writer:
    for n, df in enumerate(dfLists):
        df.to_excel(writer,sheet_name='hoja ' + str(n), index=False)
    writer.save()
