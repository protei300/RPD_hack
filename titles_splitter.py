'''
Данный модуль разбивает отсканенные титульники и сразу придает им правильные названия
'''


import pandas
from PyPDF2 import PdfFileReader, PdfFileWriter
import os
from pdf2image import convert_from_path
import pytesseract
import shutil
import re
import jellyfish

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'

PATH_TO_CORRECT_TITLES = r"D:\RPD_TEST\Титульники_2021"

PRACTICE_NAME = {}

LEARNING_CODE_PROFILE = {
    '38.04.05': 'IBA',
    '02.04.02': 'IAD',
    '09.04.01': 'IS',
    '38.03.05': 'ISiTBA',
    '09.03.04': 'RPIS',
    '09.03.03': 'PIvE',
    '02.03.02': 'IPO',
    '09.03.01': 'VMKSiS',
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
    'МБИ 2021': '38.04.05_2021',
    'МИВТ 2021': '09.04.01_2021',
    'БИ 2021': '38.03.05_2021_о',
    'БИЗ 2021': '38.03.05_2021_о-з',
    'ПИЗ 2021': '09.03.03_2021_з',
    'ПрИ 2021': '09.03.04_2021_о',
    'ПрИЗ 2021': '09.03.04_2021_з',
    'ФИИТ 2021': '02.03.02_2021_о-з',
    'МИТ 2021': '02.04.02_2021',
}


def get_disciplines_numbers(excel_code, path_to_discipline=r"D:\RPD_HACK\data\discipline_translator.xlsx"):
    try:
        translator = pandas.read_excel(path_to_discipline, sheet_name=excel_code, index_col='Номер')
    except ValueError:
        print(f"*** Не найден лист с кодом {excel_code} ")
    else:
        return translator



def recognize_name(dir_to_read=r"D:\RPD_TEST\TEMP\titles_splitted",
                   path_to_correct_titles = PATH_TO_CORRECT_TITLES,
                   title_folder = 'МБИ 2021',
                   two_sided=True,
                   clear_title_folder=True):
    '''
    Распознаем имя дисциплины в выбранном учебном плане. Это позволит определить номер титульного листа
    :param dir_to_read: Каталог откуда брать титульные листы
    :param title_folder: Название папки которая будет использоваться для сохранения титульных листов
    :return:
    '''

    print("### Начинаем распознавать титульники и создавать их с правильным названием ###")

    title_files = [os.path.join(dir_to_read, file) for file in os.listdir(dir_to_read)]
    discipline_numbers = get_disciplines_numbers(DISCIPLINE_CODES[title_folder])
    if two_sided:
        folder_to_save = os.path.join(path_to_correct_titles, title_folder)
    else:
        folder_to_save = path_to_correct_titles

    ### очистка папки с титульниками распознанными

    if not os.path.exists(folder_to_save):
        os.mkdir(folder_to_save)
    else:
        if clear_title_folder:
            for root, dirs, files in os.walk(folder_to_save):
                for f in files:
                    os.unlink(os.path.join(root, f))

    practice_number = 1

    for file in title_files:


        discipline_code = DISCIPLINE_CODES[title_folder].split('_')[0]

        doc = convert_from_path(file,
                                dpi=200,
                                poppler_path=r'D:\RPD_HACK\poppler-21.01.0\Library\bin'
                                )[0]
        path, file_name = os.path.split(file)
        fileBaseName, fileExtension = os.path.splitext(file_name)
        txt = pytesseract.image_to_string(doc, lang='rus').lower()
        splitted_text = txt.strip().split('\n')
        #print(splitted_text)

        ##### ищем слово названия дисциплины
        discipline_name = ''
        for i, line in enumerate(splitted_text):

            discipline_name = re.findall(r"рабочая программа дисци[пи]лины \"(.*)", line)
            if len(discipline_name)>0:
                discipline_name = discipline_name[0].split('\"')[0]
                #print(discipline_name)
                break



        if 'производственной практик' in txt or 'производственная практика' in txt:
            print(f"*** Создаю титульник RPP_{discipline_code}_{LEARNING_CODE_PROFILE[discipline_code]}_{practice_number} ***")
            correct_name = f"RPP_{discipline_code}_{LEARNING_CODE_PROFILE[discipline_code]}_{practice_number}.pdf"
            practice_number += 1
            shutil.copy2(file, os.path.join(path_to_correct_titles, title_folder, correct_name))
        elif 'учебной практики' in txt or 'учебная практика' in txt:

            correct_name = f"RPP_{discipline_code}_{LEARNING_CODE_PROFILE[discipline_code]}_учеб_пр.pdf"
            print(
                f"*** Создаю титульник {correct_name} ***")
            shutil.copy2(file, os.path.join(path_to_correct_titles, title_folder, correct_name))

        else:

            correct_name = f"RPD_{discipline_code}_{LEARNING_CODE_PROFILE[discipline_code]}_{discipline_name}.pdf"
            if len(discipline_name) > 0:
                found_in_table = False
                jelly_distance = []
                for row_num, row in discipline_numbers.iterrows():
                    #print(row['Дисциплина'])
                    jelly_distance.append(jellyfish.levenshtein_distance(discipline_name, row['Дисциплина'].lower()))
                    regex_name = discipline_name.replace('(','\(').replace(')','\)')
                    if  len(re.findall(f"^{regex_name}",row['Дисциплина'].lower()))>0:
                        print(f"*** Создаю титульник {row['Дисциплина']} ***")
                        correct_name = f"{row_num:02d}_RPD_{discipline_code}_{LEARNING_CODE_PROFILE[discipline_code]}.pdf"
                        shutil.copy2(file, os.path.join(folder_to_save,correct_name))
                        found_in_table = True
                        break

                if not found_in_table:
                    row_num = jelly_distance.index(min(jelly_distance)) + 1
                    print(f"*** Создаю титульник {discipline_numbers.loc[row_num, 'Дисциплина']} ***")
                    correct_name = f"{row_num:02d}_RPD_{discipline_code}_{LEARNING_CODE_PROFILE[discipline_code]}.pdf"
            else:
                print (f"!!! Не удалось распарсить дисциплину Имя файла:{file} !!!")
                correct_name = f"RPD_{discipline_code}_{LEARNING_CODE_PROFILE[discipline_code]}_{file_name.split('.')[0]}.pdf"
            shutil.copy2(file, os.path.join(folder_to_save, correct_name))

        #print("теория принятия решений" in txt.lower())
        #print(txt)

    print("### Закончили распознавать ###")



def pdf_splitter(filename, dir_to_save = r"D:\RPD_TEST\TEMP\titles_splitted", two_sided=True):
    '''
    Разбиваем PDF сканенный файл на отдельные титульники
    :param filename:
    :param dir_to_save:
    :return:
    '''

    print(f"### Приступаю к разбиению PDF файла ###")

    rpd_to_split = PdfFileReader(filename)

    if not os.path.exists(dir_to_save):
        os.mkdir(dir_to_save)
    else:
        for root, dirs, files in os.walk(dir_to_save):
            for f in files:
                os.unlink(os.path.join(root, f))

    if two_sided:
        pages = 2
    else:
        pages = 1

    for num_file, page in enumerate(range(0,rpd_to_split.getNumPages(),pages)):
        output = PdfFileWriter()
        output.addPage(rpd_to_split.getPage(page))
        if two_sided:
            output.addPage(rpd_to_split.getPage(page+1))
        with open(os.path.join(dir_to_save, f"{num_file+1}.pdf"), 'wb') as f:
            output.write(f)

    print(f"### Закончил разбиение PDF файла ###")


def merge_two_pdfs(front_dir,
                   back_dir,
                   path_to_correct_titles=PATH_TO_CORRECT_TITLES,
                   title_folder="МБИ 2021",
                   clear_title_folder=True,):
    '''
    собираем титульник итоговый из 2х страниц
    :param front_dir:
    :param back_dir:
    :param title_folder:
    :return:
    '''

    print(f"### Начинаю склеивать 2 части PDF ###")

    folder_to_save = os.path.join(path_to_correct_titles, title_folder)


    if not os.path.exists(folder_to_save):
        os.mkdir(folder_to_save)
    elif clear_title_folder:
        for root, dirs, files in os.walk(folder_to_save):
            for f in files:
                os.unlink(os.path.join(root, f))


    files_front = [file for file in os.listdir(front_dir)]
    files_back = [file for file in os.listdir(back_dir)]

    files = set.intersection(set(files_front), set(files_back))
    files_front_w_path = [os.path.join(front_dir, file) for file in files]
    files_back_w_path = [os.path.join(back_dir, file) for file in files]


    for file_front, file_back in zip(files_front_w_path, files_back_w_path):
        result_pdf = PdfFileWriter()
        front_pdf = PdfFileReader(file_front)
        back_pdf = PdfFileReader(file_back)
        result_pdf.addPage(front_pdf.getPage(0))
        result_pdf.addPage(back_pdf.getPage(0))
        file_name = file_front.split('\\')[-1]
        print(f"*** Сохраняю файл {file_name} ***")

        with open(os.path.join(folder_to_save, file_name), 'wb') as f:
            result_pdf.write(f)

    print("### Закончил клеить 2 части PDF ###")



def two_parts_splitter_merger(parts: tuple,
                              title_folder='МИВТ 2021',
                              only_merge=False,
                              clear_title_folder=True,
                              path_to_correct_titles=PATH_TO_CORRECT_TITLES):
    '''
    Разбиваем PDF состоящий из 2х файлов и склеиваем их в правильной последовательности
    :param parts: тупл из пути до 1 и 2 части
    :return:
    '''

    DIR_TO_SAVE_FRONT = r"D:\RPD_TEST\TEMP\titles_splitted_parts\front"
    DIR_TO_SAVE_BACK = r"D:\RPD_TEST\TEMP\titles_splitted_parts\back"

    DIR_TO_SAVE_FRONT_RENAMED = r"D:\RPD_TEST\TEMP\titles_splitted_parts\front_renamed"
    DIR_TO_SAVE_BACK_RENAMED = r"D:\RPD_TEST\TEMP\titles_splitted_parts\back_renamed"

    frontside, backside = parts[0], parts[1]
    paths = list(zip(parts, (DIR_TO_SAVE_FRONT, DIR_TO_SAVE_BACK), (DIR_TO_SAVE_FRONT_RENAMED, DIR_TO_SAVE_BACK_RENAMED)))

    ### Разбиваем PDF в темп директорию и распознаем содержимое

    if not only_merge:
        for path in paths:
            if not os.path.exists(path[1]):
                os.mkdir(path[1])
            else:
                for root, dirs, files in os.walk(path[1]):
                    for f in files:
                        os.unlink(os.path.join(root, f))
            pdf_splitter(path[0], path[1], two_sided=False)
            recognize_name(path[1], path_to_correct_titles=path[2], two_sided=False, title_folder=title_folder)

    ### Производим склейку 2х частей PDF

    merge_two_pdfs(DIR_TO_SAVE_FRONT_RENAMED,
                   DIR_TO_SAVE_BACK_RENAMED,
                   title_folder=title_folder,
                   clear_title_folder=clear_title_folder,
                   path_to_correct_titles=path_to_correct_titles,
                   )




def merger_splitter_controller(parts: tuple,
                               two_sided=False,
                               title_folder="МИВТ 2021",
                               path_to_correct_titles=PATH_TO_CORRECT_TITLES,
                               only_merge=False,
                               clear_title_folder=True):
    if  two_sided:
        pdf_splitter(parts[0], two_sided=True)
        recognize_name(two_sided=two_sided,
                       title_folder=title_folder,
                       clear_title_folder=clear_title_folder,
                       path_to_correct_titles=path_to_correct_titles)
    else:
        two_parts_splitter_merger(parts, title_folder=title_folder,
                                  only_merge=only_merge,
                                  path_to_correct_titles=path_to_correct_titles,
                                  clear_title_folder=clear_title_folder)


def change_page(pdf_file,
                page_to_add=None,
                pages_to_delete=None):
    '''
    Добавляем страницу, и удаляем, если надо определенную
    :param page_to_add:
    :param pdf_file:
    :return:
    '''



    result_pdf = PdfFileWriter()

    if page_to_add is not None:
        print("### Добавляю страницы ###")
        add_pdf = PdfFileReader(page_to_add)
        for page in range(add_pdf.getNumPages()):
            result_pdf.addPage(add_pdf.getPage(page))

    old_pdf = PdfFileReader(pdf_file)


    for page in range(old_pdf.getNumPages()):
        if page+1 not in pages_to_delete:
            print(f"### Добавляю страницу {page + 1} из старой PDF с учетом удаляемых")
            result_pdf.addPage(old_pdf.getPage(page))

    with open(f"{pdf_file.split('.')[0]}_m.pdf", 'wb') as f:
        result_pdf.write(f)

    print("### Закончил создание нового файла ###")





parts = (
    r"D:\RPD_TEST\Titles_to_mod\БИ\titles_BI_practice.pdf",
    #r"D:\RPD_TEST\Titles_to_mod\ФИИТ\titles_FIIT_2_2.pdf",
)

merger_splitter_controller(parts,
                           two_sided=True,
                           title_folder="БИ 2021",
                           only_merge=False,
                           clear_title_folder=False,
                           path_to_correct_titles=r"D:\RPD_TEST\Титульники_2021_практики",
                           )
'''change_page(r"D:\RPD_TEST\Titles_to_mod\БИЗ\titles_BIZ.pdf",
            page_to_add=r"D:\RPD_TEST\Titles_to_mod\БИЗ\titles_BIZ_add.pdf",
            pages_to_delete=()
            )'''

#pdf_splitter(r"D:\RPD_TEST\Титульники_2021\Титулы БЖД.pdf",)

