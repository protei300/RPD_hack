import os
import datetime
import shutil


MAIN_PATH_R = 'D:\\RPD_TEST\\FOS'
MAIN_PATH_W = 'D:\\RPD_TEST\\FOS_REGROUP'
FILTER_DATE = datetime.datetime.strptime('2021-01-31', '%Y-%m-%d')

CODE_TRANSLATOR = {
    '2017_38.03.05': 'БИ 2017',
    '2017_09.03.04': 'ПрИ 2017',
    '2019_09.03.04': 'ПрИ 2019 3++',
    '2017_09.03.01': 'ИВТЗ 2017',
    '2017_09.03.03': 'ПИЗ 2017',
    '2019_09.03.03': 'ПИЗ 2019 3++',
    '2017_02.03.02': 'ФИИТ бак 2017',
    '2019_02.03.02': 'ФИИТ бак 2019 3++',
    '2019_02.04.02': 'МИТ 2019',
    '2020_38.04.05': 'МБИ 2020',
    '2019_09.04.01': 'МИВТ 2020',
}


def main():

    if not os.path.exists(MAIN_PATH_W):
        os.mkdir(MAIN_PATH_W)

    for subdir in os.listdir(MAIN_PATH_R):
        files = [file for file in os.listdir(os.path.join(MAIN_PATH_R, subdir)) if file.endswith('.docx')]
        for file in files:
            year, form, code, name = file.split('_')
            code = code.split()[0]
            try:
                subdir_w = CODE_TRANSLATOR[f'{year}_{code}']
                fullpath_w = os.path.join(MAIN_PATH_W,subdir_w)
                if not os.path.exists(fullpath_w):
                    os.mkdir(fullpath_w)
                file_src = os.path.join(MAIN_PATH_R, subdir, file)

                mod_time = datetime.datetime.fromtimestamp(os.path.getmtime(file_src))
                if mod_time>FILTER_DATE:
                    shutil.copy2(file_src, fullpath_w)
                #print(year, form, code, name)
            except KeyError:
                continue





if __name__=='__main__':
    main()
