'''
Данная библиотека перенумировывает файлы в указанной папке
'''

import os


def reindex_files(path):
    '''
    По указанному пути переиндексирует файлы (номер в названии)
    :param path:
    :return:
    '''
    files = [file for file in os.listdir(path) if file.endswith('.pdf')]
    for number, file in enumerate(files):
        file_name = file.split('_', maxsplit=1)[1]
        new_name = f"{number+1:02d}_{file_name}"
        print(file, new_name)
        os.rename(os.path.join(path, file), os.path.join(path, new_name))

def add_numbers(path):
    '''
    Добавляем номера файлам по указанному пути
    :param path:
    :return:
    '''
    files = [file for file in os.listdir(path) if file.endswith('.pdf')]
    for number, file in enumerate(files):
        new_name = f"{number + 1:02d}_{file}"
        print(file, new_name)
        os.rename(os.path.join(path, file), os.path.join(path, new_name))

def remove_numbers(path):
    '''
    удаляем номера перед именем файла XX_имя файла.pdf
    :param path:
    :return:
    '''
    files = [file for file in os.listdir(path) if file.endswith('.pdf')]
    for file in files:
        new_name = file.split('_', maxsplit=1)[-1]
        print(file, new_name)
        os.rename(os.path.join(path, file), os.path.join(path, new_name))


if __name__ == '__main__':
    #reindex_files(r"D:\RPD_TEST\FOS_title\ФИИТ 2018")
    add_numbers(r"D:\RPD_TEST\ФОС_для_модификации\БИЗ_практики 2017")
    #remove_numbers(r"D:\RPD_TEST\TEMP\4\ПрИ_практики 2018")