import mammoth
import os
import re
import docx
import requests
import bs4

BASE_DIR = 'D:\\RPD_TEST\\Tests'

def get_test_from_docx(file):
    with open(os.path.join(BASE_DIR, file), 'rb') as docx_file:
        result = mammoth.extract_raw_text(docx_file)
        text = result.value  # The raw text
        messages = result.messages  # Any messages

    text = re.split("Упражнение \d{1,2}:", text)
    text = '\n'.join(text)
    text = re.split("Номер \d{1,2}", text)
    tests = {}

    for task in text[1:]:
        result = [line for line in task.strip().splitlines(True) if line.strip()]
        for i, line in enumerate(result):
            if 'Ответ:' in line:
                awns_num = i
        question = ''.join(result[:awns_num])
        awns = ''.join(result[awns_num + 1:])
        tests[question] = awns
    return tests


def get_test_intuit(urls):
    tests = {}

    for url in urls:
        res = requests.get(url)
        soup = bs4.BeautifulSoup(res.text, features='lxml')
        for tag in soup.find_all('div', class_='item active'):
            question = tag.find_all('pre')[0].text.strip()
            answers = ""
            for answer in tag.find_all('div'):
                anw = answer.text.strip().strip('&nbsp')

                answers += anw + '\n'

            tests[question] = answers

    return tests



def make_tests(tests, filename):


    #tests = get_test_intuit("https://eljob.ru/test/25_1")

    ##### Формируем docx файл с вопросами ############
    document = docx.Document()

    table = document.add_table(rows=1, cols=3)
    style = document.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'

    table.autofit = True
    heading_cells = table.rows[0].cells
    table.style = 'Table Grid'

    # Заголовок таблицы
    heading_cells[0].text = '№ п/п'
    heading_cells[1].text = 'Формулировка вопроса'
    heading_cells[2].text = 'Варианты ответов'


    for i, key in enumerate(tests):
        cells = table.add_row().cells
        cells[0].text = f"{i+1}."
        cells[1].text = key
        cells[2].text = tests[key]

    document.save(os.path.join(BASE_DIR, filename))

    #print(tests)


if __name__=='__main__':
    urls = [ f"https://eljob.ru/test/25_{i}" for i in range(1,29) if i not in [23,24,25,26]]
    #urls = ["https://eljob.ru/test/25_1"]
    tests = get_test_intuit(urls)
    make_tests(tests, "data_mining.docx")

