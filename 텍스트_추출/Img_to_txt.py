#이미지에 있는 텍스트 추출

import cv2
import pytesseract
import os

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'
path_ = '이미지가있는 폴더경로'
#C:/Users/ungi3/PycharmProjects/pythonProject/test_img
for i in os.listdir(path_):
    origin_img = cv2.imread(os.path.join(path_,i),cv2.IMREAD_UNCHANGED)
    cut_img = origin_img[200:234,1150:1487].copy() #전체이미지에서 원하는 부분 자르기
    text = pytesseract.image_to_string(cut_img, lang='eng') #lang='kor'
    print(i.split('.')[0])
    if len(text.split(' ')) < 3:
        print('-')
    else:
        if len(text.split(' ')[2]) == 4:
            print(text.split(' ')[2][0:2])
        else:
            print(text.split(' ')[2][0:3])

            
            
