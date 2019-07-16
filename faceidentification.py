import cv2
from openpyxl import Workbook
# create Workbook object
wb=Workbook()
# save workbook
wb.save("Record.xlsx")

face_detect = cv2.CascadeClassifier('haarcascade_frontalface_default.xml')
device=cv2.VideoCapture(0)
id = input('Enter your ID: ')
count = 0;

while(True):
    ret,img=device.read();
    gray=cv2.cvtColor(img,cv2.COLOR_BGR2GRAY)
    faces=face_detect.detectMultiScale(gray,1.3,5);
    for(x,y,w,h) in faces:
        count=count+1;
        cv2.imwrite("dataSet/data."+str(id)+"."+str(count)+".jpg",gray[y:y+h,x:x+h])
        cv2.rectangle(img,(x,y),(x+w,y+h),(0,0,255),2)
        cv2.waitKey(100);
    cv2.imshow("Face",img);
    if(count>50):
        break
device.release()
cv2.destroyAllWindows()
