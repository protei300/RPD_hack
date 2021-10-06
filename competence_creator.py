from sections_changer import RPD_CHAPTERS
import os
import docx
import mammoth
import re
import pandas as pd
import tqdm

PATH_TO_ANALYZE = (
    #r'D:\RPD_TEST\!Акк 2021 ИТОГ РПД\ПрИЗ 2019 3++',
    #r'D:\RPD_TEST\!Акк 2021 ИТОГ РПД\ФИИТ бак 2019, 2020 3++',
    #r'D:\RPD_TEST\!Акк 2021 ИТОГ РПД\ПИЗ 2019, 2020 3++',
    #r'D:\RPD_TEST\!Акк 2021 ИТОГ РПД\МИТ 2019, 2020',
    r'D:\RPD_TEST\!Акк 2021 ИТОГ РПД\МИВТ 2020',
)


FILE_TO_ANALYZE = 'Алгоритмы и анализ сложности.docx'
FILE_TO_ANALYZE = '2019 г_н__09_03_04_Программная инженерия_з_2019_plx_Информатика.docx'

SECTIONS = {
    "1": '1. Цели освоения дисциплины'.upper(),
    "2": '2. Место дисциплины в структуре ОПОП'.upper(),
    "3": '3. Компетенции обучающегося, формируемые в результате освоения дисциплины (МОДУЛЯ)'.upper(),
    "4": '4. Объем дисциплины'.upper(),
    "3.1": 'В результате освоения дисциплины обучающийся должен',
}

SECTION_1 = re.compile(r'1. ЦЕЛИ ОСВОЕНИЯ ДИСЦИПЛИНЫ|1. ОБЩИЕ ПОЛОЖЕНИЯ ПО ПРАКТИКЕ')
SECTION_2 = re.compile(r'2. МЕСТО ДИСЦИПЛИНЫ В СТРУКТУРЕ ОПОП|2. МЕСТО ПРАКТИКИ В СТРУКТУРЕ ОБРАЗОВАТЕЛЬНОЙ ПРОГРАММЫ')
SECTION_3 = re.compile(r'3. КОМПЕТЕНЦИИ ОБУЧАЮЩЕГОСЯ, ФОРМИРУЕМЫЕ В РЕЗУЛЬТАТЕ ОСВОЕНИЯ ДИСЦИПЛИНЫ \(МОДУЛЯ\)|3. ПЕРЕЧЕНЬ ПЛАНИРУЕМЫХ РЕЗУЛЬТАТОВ ОБУЧЕНИЯ')
SECTION_4 = re.compile(r'4. ОБЪЕМ ДИСЦИПЛИНЫ|4. ОБЪЕМ ПРАКТИКИ')
SECTION_3_1 = re.compile(r'В результате освоения дисциплины обучающийся должен')

COMPETENCE_REGEX = re.compile(r'(ОПК-\d{1,2}|УК-\d{1,2}|ПК-\d{1,2}):')
LEARNING_CODE_REGEX = re.compile(r'(\d{2}\.\d{2}\.\d{2})')
LEARNING_YEAR_REGEX = re.compile(r'Челябинск (\d{4}) г.')
LEARNING_FORM_REGEX = re.compile(r'(очная|заочная|очно-заочная)')
LEARNING_NAME_REGEX = re.compile(r'Рабочая программа дисциплины \(модуля\)\*\n\n|Рабочая программа практики\*\n\n')

STOP_WORDS = (
            '© ФГБОУ ВО «ЧелГУ»',
            'Рабочая программа дисциплины "',
            'Рабочая программа практики "',
            'стр. ',
            )


class Competence_creator:
    def __init__(self, path_to_file, filename):
        with open(os.path.join(path_to_file, filename), 'rb') as docx_file:
            result = mammoth.extract_raw_text(docx_file)
            text = result.value  # The raw text
            messages = result.messages  # Any messages



        title_section = SECTION_1.split(text)[0]
        learning_code = LEARNING_CODE_REGEX.findall(title_section)[0]
        learning_year = LEARNING_YEAR_REGEX.findall(title_section)[0]
        discipline_name = LEARNING_NAME_REGEX.split(title_section)[1].split('\n')[0]


        #print(learning_code, learning_year, discipline_name)

        second_third_section = SECTION_3_1.split(SECTION_4.split(SECTION_1.split(text)[1])[0])[0]
        new_text = []
        for line in second_third_section.split('\n\n'):
            if not any(stop_word in line for stop_word in STOP_WORDS):
                new_text.append(line)
        second_third_section = '\n\n'.join(new_text)

        discipline_code = second_third_section.split('Цикл (раздел) ОПОП:\n\n')[1].split('\n')[0]
        third_section = SECTION_3.split(second_third_section)[1]



        competence = COMPETENCE_REGEX.findall(third_section)
        #print(competence)


        competence_place = dict.fromkeys(competence)
        competence_number = 0
        third_section = third_section.split(2*'\n')
        third_section = [line for line in third_section if line != '' and len(re.findall(r'^для реализации',line.lower())) == 0 ]
        for i, line in enumerate(third_section):
            if competence[competence_number] in line:
                competence_place[competence[competence_number]] = i
                if competence_number == len(competence)-1: continue
                competence_number += 1

        competence_place = list(competence_place.values())
        competence_place.append(len(third_section))

        code, indicators = RPD_CHAPTERS().get_learning_plan(learning_code)

        if code == 0:
            indicators = indicators[indicators['Компетенция'].isin(list(competence))]


        indicators['Тип'] = None

        indicators.loc[indicators['Индикаторы'].str.contains(r'\.1\.'), 'Тип'] = 'Знать'
        indicators.loc[indicators['Индикаторы'].str.contains(r'\.2\.'), 'Тип'] = 'Уметь'
        indicators.loc[indicators['Индикаторы'].str.contains(r'\.3\.'), 'Тип'] = 'Владеть'




        if not indicators[indicators['Компетенция'] == 'УК-1'].empty:
            indicators = indicators.append(pd.DataFrame([[None ,'УК-1', 'Владеть']],columns=indicators.columns),ignore_index=True)

        indicators['Описание'] = None
        indicators['Код'] = discipline_code
        indicators['Название'] = discipline_name

        for i in range(len(competence_place)-1):
            competence_description = ' '.join(third_section[competence_place[i]+1:competence_place[i+1]])
            know = competence_description.split('Знать:')[1].split('Уметь:')[0]
            how = competence_description.split('Уметь:')[1].split('Владеть:')[0]
            todo = competence_description.split('Владеть:')[1]
            indicators.loc[indicators['Компетенция'] == competence[i],'Описание'] = [know,how,todo]

        self.result = indicators










#competence_creator = Competence_creator(PATH_TO_ANALYZE[-1], FILE_TO_ANALYZE)
#print(competence_creator.result[['Тип', 'Индикаторы', 'Описание']])

for path in PATH_TO_ANALYZE:
    print(f"\nProceeding to path: {path}\n")
    df = pd.DataFrame(columns=('Код', 'Название', 'Компетенция', 'Индикаторы', 'Тип', 'Описание'))
    for file in os.listdir(path):
        print(f"Reading file: {file}")
        competence_creator = Competence_creator(path, file)
        df = df.append(competence_creator.result, ignore_index=True, sort=False)

    #

    df.sort_values(['Код','Название', 'Компетенция','Индикаторы'], inplace=True)
    df['Описание'] = df['Тип'] + df['Описание']
    df = df.loc[:, df.columns != 'Тип']
    #df.set_index(['Код','Название', 'Компетенция', 'Индикаторы'], inplace=True)
    result_filename = path.split('\\')[-1]
    df.drop_duplicates(inplace=True)


    df.to_excel(f'D:\\RPD_TEST\\Компетенции\\{result_filename}.xlsx', merge_cells=True)