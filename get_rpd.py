'''
Модуль по выкачиванию аннотаций и РПД из менеджера РПД


'''


import requests
import re
import bs4
import wget
import urllib.parse
import os
import tqdm
import secret
import asyncio, aiohttp, aiofiles
import datetime
import time

#YEARS = ['2015 г.н.', '2016 г.н.', '2017 г.н.', '2018 г.н.', '2019 г.н.']
#YEARS = ['2016 г.н.', '2017 г.н.', '2019 г.н.', '2020 г.н.']
#YEARS = ['2017-2018', '2019-2020']
#YEARS = ['2017-2018']
RUP_CODES = ['020402', '090301','090302','090303','090304','090401','380305','380405',]
CHUNK = 3

#regex = re.compile(r'(020302|020402|09.?03.?01|09.?03.?02|09.?03.?03|09.?03.?04|09.?04.?01|38.?03.?05|38.?04.?05|02\.03\.02_ФИИТ_о-з_2019 \(ИИТ\)|02\.04\.02_ФИИТ_о_2019 \(ИИТ\))')
#regex = re.compile(r'(020302|09.?03.?01|09.?03.?02|09.?03.?03|09.?03.?04|38.?03.?05|02\.03\.02_ФИИТ_о-з_2019 \(ИИТ\)|02\.04\.02_ФИИТ_о_2019 \(ИИТ\))')
#regex = re.compile(r'(38.?03.?05)')
#regex = re.compile(r'(09.03.03)')
#regex = re.compile(r'02\.03\.02_ФИИТ_о-з_2019 \(ИИТ\)|02\.04\.02_ФИИТ_о_2019 \(ИИТ\)')
BASE_DIR = "D:\RPD_TEST\RPD_LOADED"

ANNOT_DIR = "D:\RPD_TEST\ANNOT"

YEARS = {
    #'2017-2018': re.compile(r'(09.?03.?04|38.?03.?05)'),
    #'2019-2020': re.compile(r'(020302|09.?03.?03|09.?03.?04|02\.03\.02_ФИИТ_о-з_2019 \(ИИТ\)|02\.04\.02_ФИИТ_о_2019 \(ИИТ\))'),
    #'2019-2020': re.compile(r'(09.?03.?04)'),
    '2021-2022': re.compile(r'(09.03.03|02.03.02 ФИИТ )'),
    #'2021-2022': re.compile(r'(02.04.02_ФИИТ|38.?03.?05|38.?04.?05)'),
    #'2018-2019': re.compile(r'(380305)'),#|090301|020302)'),
    #'2017-2018': re.compile(r'(090303|090301|020302)')
}



regex_string = []




class RPD():

    def __init__(self):
        self.login = secret.LOGIN
        self.password = secret.PASSWORD
        self.headers = {
            'Host': 'rpd.csu.ru',
            'User=Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:71.0) Gecko/20100101 Firefox/71.0',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Content-Type': 'application/x-www-form-urlencoded',
        }
        self.cookies = {
            'ASP.NET_SessionId': 'xvuykhpn4qivlsdjas3evjbx',
        }


    def get_token(self):
        url = 'http://rpd.csu.ru/Auth/Login'
        params = {
            'UserName': self.login,
            'Password': self.password,

            }

        r = requests.post(url, headers=self.headers, data=params, cookies=self.cookies,
                          allow_redirects=False)
        print (r.status_code)
        print (r.cookies['.ASPXAUTH'])
        self.cookies['.ASPXAUTH'] = r.cookies['.ASPXAUTH']

    def get_rpd_list(self):
        self.rup_iit = {}
        url = 'http://rpd.csu.ru/RPDManager/GetRUPList'
        for YEAR in YEARS.keys():
            data = {
                'year': YEAR,
            }
            r = requests.post(url, headers=self.headers, data=data, cookies=self.cookies)
            rup_codes = re.findall(r'\'keys\':\[((?:\'\d+\',?)+)\]', r.text)[0]
            rup_codes = rup_codes.replace('\'',"")
            rup_codes = rup_codes.split(',')
            soup = bs4.BeautifulSoup(r.text, features="html.parser")

            rup_names_codes = []
            for i,each_tr in enumerate(soup.find_all("tr",id=re.compile('^RUPList_DXDataRow'))):
                rup_names_codes.append((each_tr.td.string,rup_codes[i]))
            rup_names_codes.sort(key=lambda x: x[0])
            regex = YEARS[YEAR]
            self.rup_iit[YEAR] = []
            #print (rup_names_codes)
            for code in rup_names_codes:
                if  regex.findall(code[0])!=[]:
                        self.rup_iit[YEAR].append(code)
        for year in self.rup_iit.keys():
            print (f"ГОД: {year}")
            print(self.rup_iit[year])

    def get_disc_rup(self):
        url = 'http://rpd.csu.ru/RPDManager/DiscList4RUP'
        codes_url = 'http://rpd.csu.ru/RPDManager/RPDList'
        self.disc_iit={}
        for year in self.rup_iit.keys():
            for rup in self.rup_iit[year]:

                data = {
                    'rupid': rup[1],
                }
                r = requests.post(url, headers=self.headers, data=data, cookies=self.cookies)
                disc_codes = re.findall(r'\'keys\':\[((?:\'\d+\',?)+)\]', r.text)[0]
                disc_codes = disc_codes.replace('\'', "")
                disc_codes = disc_codes.split(',')

                self.codes_list = []  # Список кодов для подачи в асинхронное скачивание
                asyncio.run(self.async_find_codes(disc_codes, rup[0]))

                '''disc_codes_array = []
                for disc_code in disc_codes:
                    data = {
                    'rupRowId': disc_code,
                    'rupFileName': rup[0],

                    }
                    r = requests.post(codes_url, headers=self.headers, data=data, cookies=self.cookies)
                    disc_variant_code = re.findall(r'\'keys\':\[((?:\'\d+\',?)+)\]', r.text)
                    if disc_variant_code != []:
                        disc_variant_code = disc_variant_code[0].replace("'", '')
                    else:
                        disc_variant_code = '0'
                    disc_codes_array.append((disc_code, disc_variant_code))'''
                self.disc_iit[f"{year} - {rup[0]}"] = self.codes_list

            '''for key in self.disc_iit.keys():
                print (f"Учебный план: {key}")
                print (self.disc_iit[key])'''

    def download_rpd(self):
        for key in self.disc_iit.keys():
            check_url =  'http://rpd.csu.ru/RPDManager/RPDList'
            get_url = 'http://rpd.csu.ru/RPDPrint/ExportToWord'
            path = os.path.join(BASE_DIR,key.split(' - ')[0],key.split(' - ')[1].split('.plx')[0])
            if not os.path.isdir(path):
                os.makedirs(path)

            for i,disc in enumerate(tqdm.tqdm(self.disc_iit[key])):
                data = {
                    'rupRowId': disc,
                    'rupFileName': key.split(' - ')[1],

                }
                r = requests.post(check_url, headers=self.headers, data=data, cookies=self.cookies)
                codes = re.findall(r'\'keys\':\[((?:\'\d+\',?)+)\]', r.text)


                if codes != []:
                    codes = codes[0].replace("'",'')
                    params = {
                        'rupRowId':disc,
                        'rpdId': codes
                    }

                    r = requests.get(get_url, params = params)

                    filename = re.findall("filename\*=UTF-8''(.*)", r.headers["Content-Disposition"])[0]
                    filename = urllib.parse.unquote(filename)

                    if re.findall("plx_(.+_\.docx)", filename) != []:
                        filename = re.findall("plx_(.+_\.docx)", filename)[0]
                        filename = filename.replace('_','')
                    filename = f"{(i+1):02d} - {filename}"
                    #print (filename)

                    with open(os.path.join(path,filename),'wb') as f:
                        f.write(r.content)

    def async_download(self):
        #print (self.disc_iit.keys())
        #return
        for key in tqdm.tqdm(self.disc_iit.keys()):


            path = os.path.join(BASE_DIR,key.split(' - ')[0],key.split(' - ')[1].split('.plx')[0].replace(' ','_'))
            if not os.path.isdir(path):
                os.makedirs(path)
            #self.codes_list = []  #Список кодов для подачи в асинхронное скачивание
            #asyncio.run(self.async_find_codes(self.disc_iit[key],key.split(' - ')[1]))
            print(f"Код дисциплины {key}")
            print(f"Всего дисциплин: {len(self.disc_iit[key])}")
            '''check_url = 'http://rpd.csu.ru/RPDManager/RPDList'
            for i,disc in enumerate(self.disc_iit[key]):
                data = {
                    'rupRowId': disc,
                    'rupFileName': key.split(' - ')[1],

                }
                r = requests.post(check_url, headers=self.headers, data=data, cookies=self.cookies)
                codes = re.findall(r'\'keys\':\[((?:\'\d+\',?)+)\]', r.text)


                if codes != []:
                    codes = codes[0].replace("'",'')
                    params = (disc, codes)
                    self.codes_list.append(params)'''


            i = 0
            while True:
                print (f"===============downloading files {i}-{i+CHUNK-1}=================")
                asyncio.run(self.async_download_tasks(self.disc_iit[key][i:i+CHUNK], path))

                i += CHUNK
                if i > len(self.disc_iit[key]):
                    break
            #asyncio.run(self.async_download_tasks(self.codes_list, path))

    async def async_find_codes(self,disc_list, rupFileName):
        async def async_find_code(disc_code,rupFileName):

            check_url =  'http://rpd.csu.ru/RPDManager/RPDList'
            data = {
                'rupRowId': disc_code,
                'rupFileName': rupFileName,

            }
            async with aiohttp.ClientSession() as session:
                async with session.post(check_url, headers=self.headers, data=data, cookies=self.cookies) as r:
                    text = await r.text()
                    #codes = re.findall(r'\'keys\':\[((?:\'\d+\',?)+)\]', text)
                    codes = re.findall(r"\'rpdId':(\d+)}", text)
                    if codes != []:
                        codes = codes[-1]
                        #codes = codes[0].replace("'", '')#.split(',')[0]
                        params = (disc_code, codes)
                        self.codes_list.append(params)
                        #print(disc_code, codes)
        print(f'\n============gathering rup codes for {rupFileName} ============')
        return await asyncio.gather(*[async_find_code(disc_code,rupFileName) for disc_code in disc_list])



    async def async_download_tasks(self,codes_list,path):
        async def download_async_file(i,disc_code,path):
            #print(f'============downloading file number {i} ============')
            get_url = 'http://rpd.csu.ru/RPDPrint/ExportToWord'
            params = {
                'rupRowId': disc_code[0],
                'rpdId': disc_code[1]
            }
            async with aiohttp.ClientSession() as session:
                async with session.get(get_url, params=params, headers=self.headers, cookies=self.cookies) as r:
                    header = r.headers
                    i = 0
                    while True:
                        try:
                            filename = re.findall("filename\*=UTF-8''(.*)", header["Content-Disposition"])[0]
                            filename = urllib.parse.unquote(filename)

                            if re.findall("plx_(.+_\.docx)", filename) != []:
                                filename = re.findall("plx_(.+_\.docx)", filename)[0]
                                filename = filename.replace('_', '')
                            filename = f"{filename}"
                            # print (filename)
                            data = await r.read()
                            break
                        except Exception as e:
                            print (f"filenumber {i}, codes: {params}, error {e}, i {i}, {r.text}, {r.status}")
                            if i < 10:
                                i+=1
                            else:
                                break
                            time.sleep(2)


            path_file = os.path.join(path,filename)
            if len(path_file) > 260:
                path_file = path_file.split('.docx')[0][:-(len(path_file) - 258)] + '.docx'
            async with aiofiles.open(path_file, 'wb') as f:
                await f.write(data)

        return await asyncio.gather(*[download_async_file(i,codes, path) for i,codes in enumerate(codes_list)])

    def download_annot_rpd(self):

        get_url_1 = 'http://rpd.csu.ru/RPDPrint/PrintAnnot'
        get_annot_2 = 'http://rpd.csu.ru/FastReport.Export.axd'

        for key in self.disc_iit.keys():

            path = os.path.join(BASE_DIR, key.split(' - ')[0], key.split(' - ')[1].split('.plx')[0])
            if not os.path.isdir(path):
                os.makedirs(path)

            for i,disc in enumerate(self.disc_iit[key]):
                params = {
                    'rupRowId':disc[0],
                    'rpdId': disc[1]
                }

                r = requests.get(get_url_1, params = params)

                soup = bs4.BeautifulSoup(r.text, features="html.parser")
                id_annot = soup.find_all('script')[-1].text
                id_annot = re.findall('\/FastReport\.Export\.axd\?previewobject=(.*)\'\)', id_annot)[0]

                refresh_params = {
                    'previewobject': id_annot,
                    'refresh': 1

                }

                year_income = re.findall('^\d{4} ', key.split(' - ')[0])[0]

                r = requests.get(get_annot_2, headers=self.headers, params=refresh_params, allow_redirects=True)


                set_date_params = {
                    'object': id_annot,
                    'dialog': 2,
                    'control': 'TextBox2',
                    'event': 'onchange',
                    'data': year_income,   #Год набора
                }

                r = requests.get(get_annot_2, headers=self.headers, params=set_date_params, allow_redirects=True)

                form_annot_params = {
                    'object': id_annot,
                    'dialog': 2,
                    'control': 'btnOk1',
                    'event': 'onclick',
                    'data': '',

                }

                r = requests.get(get_annot_2, headers = self.headers, params=form_annot_params, allow_redirects=True)

                get_annot_params = {
                    'previewobject': id_annot,
                    'export_word2007': 1,
                }
                r = requests.get(get_annot_2, headers=self.headers, params=get_annot_params, allow_redirects=True)

                #print (r.text)
                #print (id_annot)
                filename = re.findall("filename\*=UTF-8''(.*)", r.headers["Content-Disposition"])[0]
                filename = urllib.parse.unquote(filename)
                print (filename)

                if re.findall("plx_(.+_\.docx)", filename) != []:
                    filename = re.findall("plx_(.+_\.docx)", filename)[0]
                    filename = filename.replace('_','')

                #print (filename)

                with open(os.path.join(path,filename),'wb') as f:
                    f.write(r.content)
                break

    def async_annot_download(self):

        for key in tqdm.tqdm(self.disc_iit.keys()):

            path = os.path.join(ANNOT_DIR,key.split(' - ')[0],key.split(' - ')[1].split('.plx')[0])
            if not os.path.isdir(path):
                os.makedirs(path)
            print (f"Качаем РУП: {key}")
            print ( f"Всего дисциплин: {len(self.disc_iit[key])}")
            '''check_url = 'http://rpd.csu.ru/RPDManager/RPDList'
            for i,disc in enumerate(self.disc_iit[key]):
                data = {
                    'rupRowId': disc,
                    'rupFileName': key.split(' - ')[1],

                }
                r = requests.post(check_url, headers=self.headers, data=data, cookies=self.cookies)
                codes = re.findall(r'\'keys\':\[((?:\'\d+\',?)+)\]', r.text)


                if codes != []:
                    codes = codes[0].replace("'",'')
                    params = (disc, codes)
                    self.codes_list.append(params)'''

            year_income = re.findall('^\d{4}', key.split(' - ')[0])[0]
            i = 0
            while True:
                print (f"===============downloading files {i}-{i+CHUNK-1}=================")
                asyncio.run(self.async_download_annot_tasks(self.disc_iit[key][i:i+CHUNK], path, year_income))

                i += CHUNK
                if i > len(self.disc_iit[key]):
                    break
            #asyncio.run(self.async_download_tasks(self.codes_list, path))

    async def async_download_annot_tasks(self,codes_list,path, year):
        async def download_async_file(i,disc_code,path, year_income):
            #print(f'============downloading file number {i} ============')
            get_url_1 = 'http://rpd.csu.ru/RPDPrint/PrintAnnot'
            get_annot_2 = 'http://rpd.csu.ru/FastReport.Export.axd'
            params = {
                'rupRowId': disc_code[0],
                'rpdId': disc_code[1]
            }

            async with aiohttp.ClientSession() as session:
                async with session.get(get_url_1, headers=self.headers, params=params) as r:

                    soup = bs4.BeautifulSoup(await r.text(), features="html.parser")
                    id_annot = soup.find_all('script')[-1].text
                    id_annot = re.findall('\/FastReport\.Export\.axd\?previewobject=(.*)\'\)', id_annot)[0]

                refresh_params = {
                    'previewobject': id_annot,
                    'refresh': 1,
                    '_': int((datetime.datetime.utcnow() - datetime.datetime(1970, 1, 1)).total_seconds() * 1000)

                }
                await session.get(get_annot_2, headers=self.headers, params=refresh_params, allow_redirects=True)

                set_date_params = {
                    'object': id_annot,
                    'dialog': 2,
                    'control': 'TextBox2',
                    'event': 'onchange',
                    'data': year_income,  # Год набора
                    '_': int((datetime.datetime.utcnow() - datetime.datetime(1970, 1, 1)).total_seconds() * 1000),
                }

                await session.get(get_annot_2, headers=self.headers, params=set_date_params, allow_redirects=True)

                form_annot_params = {
                    'object': id_annot,
                    'dialog': 2,
                    'control': 'btnOk1',
                    'event': 'onclick',
                    'data': '',
                    '_': int((datetime.datetime.utcnow() - datetime.datetime(1970, 1, 1)).total_seconds() * 1000),

                }

                await session.get(get_annot_2, headers=self.headers, params=form_annot_params, allow_redirects=True)

                get_annot_params = {
                    'previewobject': id_annot,
                    'export_word2007': 1,
                }

                async with session.get(get_annot_2, headers=self.headers, params=get_annot_params, allow_redirects=True) as r:
                    header = r.headers
                    try:
                        filename = re.findall("filename\*=UTF-8''(.*)", header["Content-Disposition"])[0]
                        filename = urllib.parse.unquote(filename)

                        if re.findall("plx_(.+_\.docx)", filename) != []:
                            filename = re.findall("plx_(.+_\.docx)", filename)[0]
                            filename = filename.replace('_', '')
                        filename = f"{filename}"
                        # print (filename)
                        data = await r.read()
                    except Exception:
                        print (f"filenumber {i}, codes: {params}")
                        return
            async with aiofiles.open(os.path.join(path, filename), 'wb') as f:
                await f.write(data)

        return await asyncio.gather(*[download_async_file(i,codes, path, year) for i,codes in enumerate(codes_list)])


if __name__ == '__main__':
    rpd = RPD()
    rpd.get_token()
    rpd.get_rpd_list()
    rpd.get_disc_rup()

    # Качаем аннотации
    rpd.async_annot_download()

    # Качаем РПД
    #rpd.async_download()