import re
import os
from tqdm import tqdm
from docx2pdf import  convert
from PyPDF2 import PdfFileReader, PdfFileWriter
import pandas as pd
import cyrtranslit
import mammoth

BASE_DIR = "D:\\RPD_TEST\\TEMP"
TEMP_DIR = "D:\\RPD_TEST\\TEMP\\4"

DOC_DIR = "D:\\RDP_TEST\\MODIFIED_SECTION"



learn_code_profile = {

    '38.04.05': 'ИБА',
    '02.04.02': 'ИАД',
    '09.04.01': 'ИС',
    '38.03.05': 'ИСиТБА',
    '09.03.04': 'РПИС',
    '09.03.03': 'ПИвЭ',
    '02.03.02': 'ИПО',
    '09.03.01': 'ВМКСиС',

}

DISCIPLINE_CODES = {
    #'БИ 2017': '38.03.05',
    #'БИЗ 2017': '38.03.05_з',
    #'ИВТЗ 2016': '09.03.01_2016_з',
    #'МБИ 2020': '38.04.05',
    #'МИВТ 2020': '09.04.01',
    #'ПИЗ 2016': '09.03.03_2016_з',
    #'ПИЗ 2019 3++': '09.03.03_2019_з',
    #'ПрИ 2017': '09.03.04_2017_о',
    #'ПрИ 2019 3++': '09.03.04_2019_о',
    #'ПрИЗ 2017': '09.03.04_2017_з',
    #'ПрИЗ 2019 3++': '09.03.04_2019_з',
    #'ФИИТ 2016': '02.03.02_2016_з',
    #'ФИИТ 2019 3++': '02.03.02_2019_о-з',
    #'МИТ 2019': '02.04.02',
    #'БИ 2018, 2019, 2020': '38.03.05_2018_о',
    #'БИЗ 2018, 2019, 2020': '38.03.05_2018_з',
    #'БИЗ 2018': '38.03.05_2018_з',
    #'БИЗ 2019-2020': '38.03.05_2018_з',
    #'ИВТЗ 2017': '09.03.01_2017_з',
    #'ИВТЗ 2018': '09.03.01_2018_з',
    #'ПИЗ 2017': '09.03.03_2017_з',
    #'ПИЗ 2018': '09.03.03_2018_з',
    #'ФИИТ 2017': '02.03.02_2017_з',
    #'ФИИТ 2018': '02.03.02_2018_з',
    #'ПрИ 2018': '09.03.04_2018_о',
    #'ПрИЗ 2018': '09.03.04_2018_з',
    #'ПИЗ 2021': '09.03.03_2021_з',
    #'ПрИЗ 2021': '09.03.04_2021_з', #+
    #'ПрИ 2021': '09.03.04_2021_о', #+
    #'ФИИТ 2021': '02.03.02_2021_о-з',
    'МИТ 2021': '02.04.02_2021', #+
    #'МБИ 2021': '38.04.05_2021',
    #'БИ 2021': '38.03.05_2021_о',
    #'БИЗ 2021': '38.03.05_2021_о-з',
    #'МИВТ 2021': '09.04.01_2021',
}


class FileScanner():

    def __init__(self):
        #self.title_path = r'D:\RPD_TEST\FOS_title'
        #self.body_path = r'D:\RPD_TEST\FOS_body'
        #self.body_path = r'D:\RPD_TEST\Add_FOS'
        self.title_path = r'D:\RPD_TEST\Титульники_2021_практики'
        self.body_path = r'D:\RPD_TEST\RPD_body_practice'
        self.save_path = r'D:\RPD_TEST\RPD_PDF_practice'
        #self.body_path = r'D:\RPD_TEST\!Акк 2021 ИТОГ РПД'

        translator_file = os.path.join('.', 'data', 'discipline_translator_practice.xlsx')
        translator = pd.read_excel(translator_file,
                                   sheet_name=None)
        self.pdf_creator = PDF_creator(self.save_path)

        #log_text = ''



        for folder in os.listdir(self.title_path):

            log_text = f'Папка {folder}\n\n'
            print(f"Обрабатываю папку {folder}")

            title_folder = os.path.join(self.title_path, folder)
            body_folder = os.path.join(self.body_path, folder)

            if os.path.isdir(title_folder) and os.path.exists(body_folder) and folder in DISCIPLINE_CODES.keys():

                if DISCIPLINE_CODES[folder] in translator.keys():
                    translator_df = translator[DISCIPLINE_CODES[folder]]
                    title_files = [file for file in os.listdir(title_folder) if file.endswith('.pdf')]
                    body_files_dict = self.create_body_names_dict(body_folder)
                    #print(body_files_dict.keys())

                    for file in title_files:

                        file_code = file.split('_')[0]
                        discipline_series = translator_df.loc[translator_df['Номер'] == int(file_code), 'Дисциплина']

                        if len(discipline_series) > 0:
                            discipline_name = discipline_series.values[0].lower()
                            if discipline_name in body_files_dict.keys():
                                print(f"*** Обрабатываю дисциплину {discipline_name} в папке {folder} ***")
                                '''result_pdf_filename = re.findall(r"(\d{1,2}_(?:RPD_|RPP_|FOS_){0,1}\d{2}\.\d{2}\.\d{2}_[A-Za-z]*)[_\.\s]",
                                                                 file)[0]'''
                                result_pdf_filename = file.split('.pdf')[0]
                                code, doc_type,  learn_code, profile = result_pdf_filename.split('_', maxsplit=3)
                                #print(profile)
                                if doc_type == 'FOS':
                                    discipline_name_short = [word[0:3] for word in discipline_name.split(' ')]
                                    result_pdf_filename = f"{int(code):02d}_ФОС_{learn_code}_{learn_code_profile[learn_code]}_{' '.join(discipline_name_short)}"
                                elif doc_type == 'RPD':
                                    result_pdf_filename += '_' + cyrtranslit.to_latin(discipline_name, 'ru')
                                elif doc_type == 'RPP':
                                    result_pdf_filename = result_pdf_filename.split('_', maxsplit=4)
                                    tail = cyrtranslit.to_latin(result_pdf_filename[-1], 'ru')
                                    result_pdf_filename = f"{doc_type}_{learn_code}_{profile.split('_', maxsplit=1)[0]}_{tail}"

                                else:
                                    if 'гиа' in profile.lower():
                                        result_pdf_filename = f"{doc_type}_{learn_code_profile[doc_type]}_ГИА"
                                    elif 'учеб' in profile.lower():
                                        discipline_name_short = [word[:1] for word in re.split('[\s-]', discipline_name)]
                                        result_pdf_filename = f"{doc_type}_{learn_code_profile[doc_type]}_УП({''.join(discipline_name_short)})"
                                    else:
                                        discipline_name_short = [word[:1] for word in re.split('[\s-]', discipline_name)]
                                        result_pdf_filename = f"{doc_type}_{learn_code_profile[doc_type]}_ПП({''.join(discipline_name_short)})"

                                result_pdf_filename += '.pdf'
                                title_filename = os.path.join(title_folder, file)
                                body_filename = body_files_dict[discipline_name]
                                #print(title_filename, body_filename, folder, result_pdf_filename)
                                self.pdf_creator.create_pdf_from_docx(title_filename, body_filename, folder, result_pdf_filename)
                            else:
                                log_text += f'Дисциплина {discipline_name} не найдена в папке с телами\n'
                                print(f"!!! Дисциплина {discipline_name} не найдена в папке с телами !!!")

                        else:
                            log_text += f'Дисциплина {file} по коду в таблице транслятора не найдена\n'



                else:
                    log_text += f'Такого {DISCIPLINE_CODES[folder]} нет в таблице транслятора\n'



                with open(os.path.join(self.save_path,folder,'log.txt'), 'w+') as f:
                    f.write(log_text)

    def create_body_names_dict(self, path="D:\RPD_TEST\RPD_body\ПИЗ 2021"):
        '''
        Создаем словарь реальное название дисциплины - имя файла
        Определять будем по содержимому титульного листа
        :return: возвращаем словарь ключ - дисциплина, содержимое имя файла
        '''

        print ("### Начинаю создание словаря тел РПД ###")

        files = [os.path.join(path, file) for file in os.listdir(path) if file.endswith('.docx')]

        body_names_dict = {}

        for file in files:

            print(f"*** Обрабатываю файл {os.path.split(file)[1]} ***")
            try:
                doc_txt = mammoth.extract_raw_text(file).value
            except Exception:
                continue
            else:
                doc_txt = doc_txt.strip().split('Направление подготовки (специальность)')[0].lower()
                key = re.search("(?:рабочая программа дисциплины \(модуля\)|рабочая программа практики)\*\n\n(.*)(\n|$)", doc_txt.lower())

                if key is not None:
                    body_names_dict[key.group(1).strip()] = file
                else:
                    print(f"!!! Не найден паттерн в файле {os.path.split(file)[1]} !!!")
                    print(doc_txt)

            #print(body_names_dict)

        print("### Закончил создание словаря тел РПД ###")

        return body_names_dict




    def make_discpline_name(self, folder):
        files = os.listdir(folder)
        return [file.lower().split('.')[-2].split('_')[-1] for file in files]


class PDF_creator():
    '''
    Класс по склеиванию PDF титульника и docx тела документа
    '''
    def __init__(self, folder_to_save):
        self.folder_to_save = folder_to_save

    def remove_element(self, element):
        element.getparent().remove(element)

    def create_pdf_from_docx(self,title_filename, body_filename, target_folder, target_filename):
        '''
        Создаем docx без титульных листов
        :return:
        '''


        #конвертируем в pdf
        convert(body_filename, os.path.join(TEMP_DIR,target_filename))

        #прилепляем титульник
        title = PdfFileReader(title_filename)
        body = PdfFileReader(os.path.join(TEMP_DIR, target_filename))
        #body = PdfFileReader(body_filename)
        output = PdfFileWriter()
        for page_number in range(title.getNumPages()):
            output.addPage(title.getPage(page_number))
        for page_number in range(2,body.getNumPages()):
            output.addPage(body.getPage(page_number))

        if not os.path.exists(os.path.join(self.folder_to_save, target_folder)):
            os.makedirs(os.path.join(self.folder_to_save, target_folder))

        with open(os.path.join(self.folder_to_save, target_folder, target_filename), 'wb') as f:
            output.write(f)

        return 0


class TitleNamer:
    '''
    Класс по добавлению _RPD_ в название титульников в формате pdf
    '''


    def __init__(self, folder):
        self.main_folder = folder



    def make_rpd_names(self):
        for folder in os.listdir(self.main_folder):
            folder_medium = os.path.join(self.main_folder, folder)
            if not os.path.isdir(folder_medium):
                continue
            for file in os.listdir(folder_medium):
                if file.endswith('.pdf') and "RPD" not in file:
                    print(file)
                    splitted_name = file.split('_', maxsplit=1)
                    new_name = splitted_name[0] + "_RPD_" + splitted_name[1]
                    os.rename(os.path.join(folder_medium, file),
                              os.path.join(folder_medium, new_name))
                    print(new_name)

    def file_lister(self):
        ''' Функция перебора файлов внутри каталогов '''
        for folder in os.listdir(self.main_folder):
            folder_medium = os.path.join(self.main_folder, folder)
            if not os.path.isdir(folder_medium):
                continue
            files_list = [file for file in os.listdir(folder_medium) if file.endswith('.pdf')]
            for file in files_list:
                try:
                    os.rename(os.path.join(folder_medium,file), os.path.join(folder_medium, self.make_short_names(file)))
                except FileExistsError:
                    print(os.path.join(folder_medium,file))



    def make_short_names(self, filename):
        ''' Укорачиваем имя файлам '''

        filename_name_parts = filename.split('.pdf')[0].split('_')
        new_filename = ''
        new_filename_name = [word[0:3] for word in filename_name_parts[-1].split(' ')]
        new_filename_name = ' '.join(new_filename_name)
        new_filename = '_'.join(filename_name_parts[:-1]) + '_'+ new_filename_name + '.pdf'
        return new_filename



#TitleNamer(r"D:\РПД\!Сканы титульных листов РПД 2021").make_rpd_names()



#TitleNamer('').make_short_names('1_RPD_38.03.05_ISiTBA_Inostrannyj jazyk.pdf')
#TitleNamer('D:\RPD_TEST\!Акк готовые ФОС').file_lister()

FileScanner()


#pdf_creator = PDF_creator()
#pdf_creator.create_pdf_from_docx('Базы и хранилища данных.docx')