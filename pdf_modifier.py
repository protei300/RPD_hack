'''
В этой библиотеке будут храниться классы по преобразованию PDF
'''

import re
from datetime import datetime
import fitz
import cv2
import os
import cyrtranslit
import img2pdf
from pdf2image import convert_from_path
from PyPDF2 import PdfFileReader, PdfFileWriter


FOLDERS = [
    #'МИТ 2019',
    #'МИТ 2020',
    #'ПИЗ 2018',
    #'ПИЗ 2019',
    #'ПрИ 2019',
    #'ПрИЗ 2019',
    #'ПрИ 2017',
    #'БИ 2017',
    #'БИ_практики 2018',
    #'БИЗ_практики 2018',
    #'БИЗ_практики 2017',
    #'ПрИ_практики 2018',
    #'БИ 2018, 2019, 2020',
    #'ПрИ 2018',
    #'ФИИТ 2019',
    #'ФИИТ 2018',
    #'ИВТЗ 2018',
    #'ПрИЗ 2021',
    #'МБИ 2021'
    'БИ 2021'
    #'На изменение'
]

Specific_numbers = [2,3,4,5 ]
#Specific_numbers = [13,14 ]
#Specific_numbers = [16, 17, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 54, 55]

TEST = False
use_specific_numbers = True
SECOND_PAGE = False #### изменяем 2 страницу
CHANGE_YEAR = True #### Меняем год набора
CHANGE_LEARNING_FORM = False #### Меняем форму набора


COORDS = 770


class PDFTitleModifier:
    '''
    в данном классе работаем с титульниками в формате pdf
    '''


    def __init__(self, filename_to_read, second_page=False):

        self.TEMP_DIR = r'D:\RPD_TEST\TEMP\4'
        self.TEMP_IMG_FILE = os.path.join(self.TEMP_DIR, '1.png')
        self.TEMP_PDF_FILE = os.path.join(self.TEMP_DIR, '1.pdf')
        self.second_page = second_page


        self.main_folder = filename_to_read
        self.filename = filename_to_read.split('\\')[-1].split('.pdf')[0]
        self.short_filename = self.filename.split('_')[0]
        #print(self.filename)
        self.temp_dir = ''
        pics = convert_from_path(filename_to_read,
                                dpi=96,
                                poppler_path=r'D:\RPD_HACK\poppler-21.01.0\Library\bin')[int(second_page)]

        self.file_pics = []

        pics.save(os.path.join(self.TEMP_DIR, f'{self.short_filename}.png'))

        '''for i, pic in enumerate(pics):
            pic.save(os.path.join(self.TEMP_DIR, f'{i:02d}_{self.filename}.png'))
            self.file_pics.append(f'{i:02d}_{self.filename}.png')'''
        #pdf = fitz.open(filename_to_read)
        #page = pdf.loadPage(0)
        #pic = page.getPixmap()
        #pic.writePNG(os.path.join(self.TEMP_DIR, f'{self.filename}.png'))

    def change_colontitul(self):
        '''
        Функция меняет на картинке года
        :param path_to_write:
        :return:
        '''

        # Load images, grayscale, Gaussian blur, Otsu's threshold
        for file in self.file_pics:
            original = cv2.imread(os.path.join(self.TEMP_DIR, file))
            image = cv2.imread(r'D:\RPD_TEST\Ksenya\colontitul.png')

            # print(original.shape)
            dsize = (original.shape[1], image.shape[0])

            # resize image
            image = cv2.resize(image, dsize, interpolation=cv2.INTER_AREA)

            gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
            blur = cv2.GaussianBlur(gray, (3, 3), 0)
            thresh = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]

            # Find contours, filter using contour approximation + area, then extract
            # ROI using Numpy slicing and replace into original image
            cnts = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            cnts = cnts[0] if len(cnts) == 2 else cnts[1]
            # print(cnts)
            for c in cnts:
                peri = cv2.arcLength(c, True)
                approx = cv2.approxPolyDP(c, 0.015 * peri, True)
                area = cv2.contourArea(c)
                # print(len(approx),area)
                if len(approx) == 4 and area > 1000:
                    x, y, w, h = cv2.boundingRect(c)
                    ROI = image[y:y + h, x:x + w]
                    original[y + COORDS:y + COORDS + h, x:x + w] = ROI

            # cv2.imshow('thresh', thresh)
            # cv2.imshow('ROI', ROI)
            # cv2.imshow('original', original)
            # cv2.waitKey()

            # path_to_write = cyrtranslit.to_latin(path_to_write, 'ru')

            os.remove(os.path.join(self.TEMP_DIR, file))
            cv2.imwrite(os.path.join(self.TEMP_DIR, file), original)


    def change_year_second_page(self):
        '''
        Функция меняет дату протокола на 2 странице
        :param path_to_write:
        :return:
        '''

        # Load images, grayscale, Gaussian blur, Otsu's threshold
        original = cv2.imread(os.path.join(self.TEMP_DIR,self.short_filename + '.png'))
        #original = cv2.imread(self.TEMP_IMG_FILE)
        image = cv2.imread(r'D:\RPD_TEST\modified_title\2_page_date.png')

        #print(original.shape)
        dsize = (original.shape[1], image.shape[0])

        # resize image
        image = cv2.resize(image, dsize, interpolation=cv2.INTER_AREA)


        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        blur = cv2.GaussianBlur(gray, (3, 3), 0)
        thresh = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]

        # Find contours, filter using contour approximation + area, then extract
        # ROI using Numpy slicing and replace into original image
        cnts = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        cnts = cnts[0] if len(cnts) == 2 else cnts[1]
        #print(cnts)
        for c in cnts:
            peri = cv2.arcLength(c, True)
            approx = cv2.approxPolyDP(c, 0.015 * peri, True)
            area = cv2.contourArea(c)
            #print(len(approx),area)
            if len(approx) == 4 and area > 1000:
                x, y, w, h = cv2.boundingRect(c)
                ROI = image[y:y + h, x:x + w]
                original[y+COORDS:y+COORDS + h, x:x + w] = ROI
                original[y+COORDS+320:y+COORDS+320 + h, x:x + w] = ROI

        #cv2.imshow('thresh', thresh)
        #cv2.imshow('ROI', ROI)
        #cv2.imshow('original', original)
        #cv2.waitKey()

        #path_to_write = cyrtranslit.to_latin(path_to_write, 'ru')


        os.remove(os.path.join(self.TEMP_DIR,self.short_filename + '.png'))
        cv2.imwrite(self.TEMP_IMG_FILE, original)

    def change_year(self):
        '''
        Функция меняет на картинке года
        :param path_to_write:
        :return:
        '''

        # Load images, grayscale, Gaussian blur, Otsu's threshold
        original = cv2.imread(os.path.join(self.TEMP_DIR,self.short_filename + '.png'))
        #original = cv2.imread(self.TEMP_IMG_FILE)
        image = cv2.imread(r'D:\RPD_TEST\modified_title\year_collection.png')

        #print(original.shape)
        dsize = (original.shape[1], image.shape[0])

        # resize image
        image = cv2.resize(image, dsize, interpolation=cv2.INTER_AREA)


        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        blur = cv2.GaussianBlur(gray, (3, 3), 0)
        thresh = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]

        # Find contours, filter using contour approximation + area, then extract
        # ROI using Numpy slicing and replace into original image
        cnts = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        cnts = cnts[0] if len(cnts) == 2 else cnts[1]
        #print(cnts)
        for c in cnts:
            peri = cv2.arcLength(c, True)
            approx = cv2.approxPolyDP(c, 0.015 * peri, True)
            area = cv2.contourArea(c)
            #print(len(approx),area)
            if len(approx) == 4 and area > 1000:
                x, y, w, h = cv2.boundingRect(c)
                ROI = image[y:y + h, x:x + w]
                original[y+COORDS:y+COORDS + h, x:x + w] = ROI

        #cv2.imshow('thresh', thresh)
        #cv2.imshow('ROI', ROI)
        #cv2.imshow('original', original)
        #cv2.waitKey()

        #path_to_write = cyrtranslit.to_latin(path_to_write, 'ru')


        os.remove(os.path.join(self.TEMP_DIR,self.short_filename + '.png'))
        cv2.imwrite(self.TEMP_IMG_FILE, original)

    def change_learning_form(self):
        '''
        Функция меняет на картинке очку на заочку
        :param path_to_write:
        :return:
        '''

        # Load images, grayscale, Gaussian blur, Otsu's threshold
        original = cv2.imread(os.path.join(self.TEMP_DIR,self.short_filename + '.png'))
        image = cv2.imread(r'D:\RPD_TEST\modified_title\ochno_zaochno.png')

        #print(original.shape)
        dsize = (original.shape[1], image.shape[0])

        # resize image
        image = cv2.resize(image, dsize, interpolation=cv2.INTER_AREA)


        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        blur = cv2.GaussianBlur(gray, (3, 3), 0)
        thresh = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]

        # Find contours, filter using contour approximation + area, then extract
        # ROI using Numpy slicing and replace into original image
        cnts = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        cnts = cnts[0] if len(cnts) == 2 else cnts[1]
        #print(cnts)
        for c in cnts:
            peri = cv2.arcLength(c, True)
            approx = cv2.approxPolyDP(c, 0.015 * peri, True)
            area = cv2.contourArea(c)
            #print(len(approx),area)
            if len(approx) == 4 and area > 1000:
                x, y, w, h = cv2.boundingRect(c)
                ROI = image[y:y + h, x:x + w]
                original[y+COORDS:y+COORDS + h, x:x + w] = ROI

        #cv2.imshow('thresh', thresh)
        #cv2.imshow('ROI', ROI)
        #cv2.imshow('original', original)
        #cv2.waitKey()

        #path_to_write = cyrtranslit.to_latin(path_to_write, 'ru')


        os.remove(os.path.join(self.TEMP_DIR,self.short_filename + '.png'))
        cv2.imwrite(self.TEMP_IMG_FILE, original)

    def convert_to_pdf(self, file_to_read):
        '''
        Перегоняем png файлы в pdf
        '''

        for file in self.file_pics:
            with open(os.path.join(self.TEMP_DIR, "1", f"{file.split('.png')[0]}.pdf"), 'wb') as f:
                f.write(img2pdf.convert(os.path.join(self.TEMP_DIR, file)))

        title = PdfFileReader(file_to_read)

        body = []
        for file in os.listdir(r"D:\RPD_TEST\TEMP\4\1"):
            print(file)
            body.append(PdfFileReader(os.path.join(r"D:\RPD_TEST\TEMP\4\1", file)))
        output = PdfFileWriter()
        output.addPage(title.getPage(0))
        output.addPage(title.getPage(1))
        for body_pdfs in body:
            output.addPage(body_pdfs.getPage(0))

        with open(os.path.join(self.TEMP_DIR, "2", f"{self.filename}.pdf"), 'wb') as f:
            output.write(f)




    def make_new_pdf(self, path_to_save, filename):

        with open(self.TEMP_PDF_FILE, 'wb') as f:
            f.write(img2pdf.convert(self.TEMP_IMG_FILE))

        if not os.path.exists(os.path.join(self.TEMP_DIR,path_to_save)):
           os.mkdir(os.path.join(self.TEMP_DIR,path_to_save))

        title = PdfFileReader(self.TEMP_PDF_FILE)
        body = PdfFileReader(self.main_folder)
        body_num_pages = body.getNumPages()
        output = PdfFileWriter()
        if not self.second_page:
            output.addPage(title.getPage(0))
            for i in range(1,body_num_pages):
                output.addPage(body.getPage(i))
        else:
            output.addPage(body.getPage(0))
            output.addPage(title.getPage(0))

        with open(os.path.join(self.TEMP_DIR, path_to_save, filename), 'wb') as f:
            output.write(f)




def main():
    main_folder = r'D:\RPD_TEST\Titles_to_mod\БИ'
    for folder in os.listdir(main_folder):
        folder_medium = os.path.join(main_folder, folder)
        if not os.path.isdir(folder_medium) or folder not in FOLDERS:
            continue
        files = [file for file in os.listdir(folder_medium) if file.endswith('pdf')]
        for file in files:
            if use_specific_numbers and int(file.split('_')[0]) in Specific_numbers or not use_specific_numbers:

                if file.endswith('.pdf'):
                    pdf = PDFTitleModifier(os.path.join(folder_medium, file), second_page=SECOND_PAGE)

                    if SECOND_PAGE:
                        pdf.change_year_second_page()
                    elif CHANGE_YEAR:
                        pdf.change_year()
                    elif CHANGE_LEARNING_FORM:
                        pdf.change_learning_form()
                    else:
                        return
                    if  TEST: return
                    pdf.make_new_pdf(folder, file)



def change_1_file(path):
    pdf = PDFTitleModifier(path, second_page=SECOND_PAGE)
    if SECOND_PAGE:
        pdf.change_year_second_page()
    elif CHANGE_YEAR:
        pdf.change_year()
    elif CHANGE_LEARNING_FORM:
        pdf.change_learning_form()
    else:
        return
    if TEST: return
    folder, filename = os.path.split(path)
    filename = f"{filename.split('.pdf')[0]}_mod.pdf"
    pdf.make_new_pdf(folder, filename)


def ksenya_changer():
    main_folder = r'D:\RPD_TEST\Ksenya'
    files = [file for file in os.listdir(main_folder) if file.endswith('.pdf')]

    for file in files:
        pdf = PDFTitleModifier(os.path.join(main_folder, file))
        pdf.change_colontitul()
        pdf.convert_to_pdf(os.path.join(main_folder,file))


if __name__=='__main__':
    #main()
    change_1_file(r"D:\RPD_TEST\Titles_to_mod\БИЗ\titles_BIZ_add.pdf")