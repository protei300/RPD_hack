'''
В данном модуле представлена библиотека по копированию файлов содержащих ключевые слова из 1 папки в другую,
с учетом вложенности
'''

import os
import shutil



def copy_files_contain_word(words, path_from, path_to):
    '''

    :param words: набор слов, которое может быть в названии файла
    :param path_from: корневой путь, откуда копировать
    :param path_to: корневой путь куда копировать
    :return:
    '''

    dirs = [dir for dir in os.listdir(path_from) if os.path.isdir(os.path.join(path_from,dir))
                                                                  and not os.path.exists(os.path.join(path_to, dir))]

    for dir in dirs:
        os.mkdir(os.path.join(path_to, dir))


    for file in os.listdir(path_from):
        current_path = os.path.join(path_from, file)
        if os.path.isfile(current_path):
            if any(word in file.lower() for word in words):
                if not os.path.isfile(os.path.join(path_to, file)):
                    shutil.copy2(current_path, path_to)
                print(f"Is practice {file}")
        elif os.path.isdir(current_path):
            print(f"Is dir {current_path} ")
            copy_files_contain_word(
                words,
                path_from = current_path,
                path_to = os.path.join(path_to, file),

                                    )

    return




if __name__=='__main__':
    copy_files_contain_word(['_гиа', '_преддипломная', '_практика', '_научно-исследовательская', '_нир', '_технологическая'],
                            'D:\РПД\!Акк 2021 ИТОГ ФОС\на аккредитацию',
                            'D:\\RPD_TEST\\FOS_body_practice',
                            )



