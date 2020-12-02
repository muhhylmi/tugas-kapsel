import cv2
import pandas as pd
from pandas import DataFrame
import os

vidcap = cv2.VideoCapture('video/Berita1.mkv')

try:
    if not os.path.exists('data'):
        os.makedirs('data')
except OSError:
    print ('Error: Creating directory of data')

def getFrame(sec):
    vidcap.set(cv2.CAP_PROP_POS_MSEC, sec*1000) 
    hasFrames,image = vidcap.read()
    if hasFrames:
        cv2.imwrite("./data/frame"+str(count)+".jpg", image)     # save frame as JPG file
        print ('Creating frame '+ str(count))
    return hasFrames

sec = 0
frameRate = 0.5                    #//it will capture image in each 0.5 second
count=0
success = getFrame(sec)
while success:
    count = count + 1
    sec = sec + frameRate
    sec = round(sec, 2)
    success = getFrame(sec)

y0=434
x0=0
h0=180
w0=1280

x1=986
y1=61
h1=50
w1=211

for i in range (3):
    img0 = cv2.imread('data/frame'+str(i)+'.jpg',1)
    img1 = cv2.imread('data/frame'+str(i+1)+'.jpg',1)
    img2 = cv2.imread('data/frame3.jpg',1)
    crop_Logo = img0[y1:y1+h1, x1:x1+w1]
    crop_Logo1= img1[y1:y1+h1, x1:x1+w1]
    crop_Logo2= img2[y1:y1+h1, x1:x1+w1]
    
    b0,g0,r0 = cv2.split(crop_Logo)
    b1,g1,r1 = cv2.split(crop_Logo1)
    
    cv2.imwrite("./db/logo3.jpg",crop_Logo2)
    cv2.imwrite("./db/logo"+str(i)+".jpg", crop_Logo)
    with pd.ExcelWriter('ecxel/Logo'+str(i)+'.xlsx')as writer:
       pd.DataFrame(r0).to_excel(writer, sheet_name="r"+str(i)+"", index=False)
       pd.DataFrame(g0).to_excel(writer, sheet_name="g"+str(i)+"", index=False)
       pd.DataFrame(b0).to_excel(writer, sheet_name="b"+str(i)+"", index=False)
       pd.DataFrame(r1).to_excel(writer, sheet_name="r"+str(i+1)+"", index=False)
       pd.DataFrame(g1).to_excel(writer, sheet_name="g"+str(i+1)+"", index=False)
       pd.DataFrame(b1).to_excel(writer, sheet_name="b"+str(i+1)+"", index=False)
       pd.DataFrame(r0-r1).to_excel(writer, sheet_name="r"+str(i)+"-r"+str(i+1)+"", index=False)
       pd.DataFrame(g0-g1).to_excel(writer, sheet_name="g"+str(i)+"-g"+str(i+1)+"", index=False)
       pd.DataFrame(b0-b1).to_excel(writer, sheet_name="b"+str(i)+"-b"+str(i+1)+"", index=False)
       
cv2.waitKey(0)
cv2.destroyAllWindows()