# This program is written by Abubakr Shafique (abubakr.shafique@gmail.com)

# 원본 사이트 주소 -> https://github.com/abubakr-shafique/Dicom_to_Image-Python/blob/master/Dcm_to_Img_Single.py
# 원본은 dicom파일을 jpg로 변경해주는 코드가 있고 여기서 필요에따라 변경한 코드
import cv2 as cv
import numpy as np
import pydicom as PDCM
import os
import openpyxl

def Dicom_to_Image(Path):
    DCM_Img = PDCM.read_file(Path)

    rows = DCM_Img.get(0x00280010).value #Get number of rows from tag (0028, 0010)
    cols = DCM_Img.get(0x00280011).value #Get number of cols from tag (0028, 0011)

    #Instance_Number = int(DCM_Img.get(0x00200013).value) #Get actual slice instance number from tag (0020, 0013)

    Window_Center = int(DCM_Img.get(0x00281050).value) #Get window center from tag (0028, 1050)
    Window_Width = int(DCM_Img.get(0x00281051).value) #Get window width from tag (0028, 1051)

    Window_Max = int(Window_Center + Window_Width / 2)
    Window_Min = int(Window_Center - Window_Width / 2)

    if (DCM_Img.get(0x00281052) is None):
        Rescale_Intercept = 0
    else:
        Rescale_Intercept = int(DCM_Img.get(0x00281052).value)

    if (DCM_Img.get(0x00281053) is None):
        Rescale_Slope = 1
    else:
        Rescale_Slope = int(DCM_Img.get(0x00281053).value)

    New_Img = np.zeros((rows, cols), np.uint8)
    Pixels = DCM_Img.pixel_array

    for i in range(0, rows):
        for j in range(0, cols):
            Pix_Val = Pixels[i][j]
            Rescale_Pix_Val = Pix_Val * Rescale_Slope + Rescale_Intercept

            if (Rescale_Pix_Val > Window_Max): #if intensity is greater than max window
                New_Img[i][j] = 255
            elif (Rescale_Pix_Val < Window_Min): #if intensity is less than min window
                New_Img[i][j] = 0
            else:
                New_Img[i][j] = int(((Rescale_Pix_Val - Window_Min) / (Window_Max - Window_Min)) * 255) #Normalize the intensities

    return New_Img

def sheetwrite(path,img):

    path_ = os.path.join(path,img)
    ds = PDCM.read_file(path_)
    PatientID = ds.get(0x00100020).value
    dcm_img_name = 'n_img'
    answer = 'n_ans'
    return ds.StudyDate,PatientID,dcm_img_name,answer

def main():
    in_path_ = 'DCM파일이있는폴더경로'
    f_name = os.listdir(in_path_)
    wb = openpyxl.Workbook()
    sheet = wb.active
    row_num = 1
    for dcm_img in f_name:
        Input_Image = os.path.join(in_path_,dcm_img)

        study_date,patientID,dcm_img_name,answer = sheetwrite(in_path_,dcm_img)
        ds = PDCM.read_file(Input_Image)
        Output_Image = Dicom_to_Image(Input_Image)
        cv.imwrite('n_ans'+ str(row_num) + '.jpg', Output_Image)  #jpg로 바꾼 이미지이름 지정

        sheet['A'+ str(row_num)] = dcm_img[:dcm_img.find('.')]
        sheet['B'+ str(row_num)] = patientID
        sheet['C'+ str(row_num)] = dcm_img[1:8]
        sheet['D'+ str(row_num)] = study_date
        sheet['E'+ str(row_num)] = dcm_img_name + str(row_num)
        sheet['F'+ str(row_num)] = answer + str(row_num)
        ds.PatientID = dcm_img[1:8]
        a = '저장할폴더경로'+ dcm_img_name + str(row_num) +'.dcm' #dicom파일 읽고 변경한 파일
        ds.save_as(a)
        row_num += 1
        print(str(row_num) +" "+dcm_img+" 완료")

    wb.save('test.xlsx')

if __name__ == "__main__":
    main()