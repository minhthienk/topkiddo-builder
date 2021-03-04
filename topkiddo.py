import os
import sys
import json
import requests
import random,string
from requests_toolbelt.multipart.encoder import MultipartEncoder
import re
import time


import pandas as pd
from pathlib import Path


#todo: kiem tra khoang trang, case sensitive
class Excel:
    # Contructor
    def __init__(self, path, sheet_names):
        self.sheets = self.load_dtb(path, sheet_names) # a dict to contain all sheets data frame
    # load dtb from excel or pickle files
    @staticmethod
    def load_dtb(path, sheet_names):
        # create data frames from pickle files if not create pickle files
        sheets = {}
        for sheet_name in sheet_names:
            print('read excel files: ', path, ' - ',sheet_name)
            sheets[sheet_name] = pd.read_excel(path, 
                                               sheet_name=sheet_name, 
                                               header=0, 
                                               na_values='#### not defined ###', 
                                               keep_default_na=False)
        return sheets


def input_excel_database(sheet):
    sheets = [sheet]
    excel_path = 'TOPKIDDO\\Topkiddo.xlsx'
    df = Excel(excel_path, sheets).sheets[sheet]
    # strip all strings from excel database
    df.replace(r'(^\s+|\s+$)', '', regex=True, inplace=True)
    df.replace(r'\s+', ' ', regex=True, inplace=True)
    return df


def get_word_data(df, word):
    word = word.strip()
    root_row = df[df['Word'].str.match('^' + word + '$', na=False, case=False)].index.values

    if len(root_row)==0:
        return []

    audio = df.iloc[root_row + 5, 1].values[0]
    video = df.iloc[root_row + 6, 16].values[0]

    data = []
    for step in range(1,8+1):
        root_step_column = step*2
        step_data = {
            'content':df.iloc[root_row + 1, root_step_column].values[0],
            'position':df.iloc[root_row + 2, root_step_column].values[0],
            'language':df.iloc[root_row + 3, root_step_column].values[0],
            'tags':df.iloc[root_row + 4, root_step_column].values[0],
            'audio':df.iloc[root_row + 5, 1].values[0],
            'image':df.iloc[root_row + 6, root_step_column].values[0],
            'video':df.iloc[root_row + 6, 16].values[0]
        }
        if step!=8:
            step_data['video']=''
        else:
            step_data['audio']=''
            step_data['image']=''

        # image for step 1
        if step==1:
            step_data['image']=df.iloc[root_row + 6, root_step_column-1].values[0]
        
        # correct file paths:
        step_data['image'] = step_data['image'].replace("'", '_')
        step_data['audio'] = step_data['audio'].replace("'", '_')
        step_data['video'] = step_data['video'].replace("'", '_')

        data.append(step_data)


    return data



def get_sentence_data(df, sentence):
    sentence = sentence.strip()

    # get root row and root column
    for label, content in df.items():
        root_row = df[df[label].str.match('^' + sentence + '$', na=False, case=False)].index.values
        if len(root_row)!=0:
            root_column = df.columns.get_loc(label)
            break

    # check if not found => return []
    if len(root_row)==0:
        return []

    data = {
        'content':df.iloc[root_row, root_column].values[0],
        'position':df.iloc[root_row + 1, root_column].values[0],
        'language':df.iloc[root_row + 2, root_column].values[0],
        'tags':df.iloc[root_row + 3, root_column].values[0],
        'audio':df.iloc[root_row + 4, root_column].values[0],
        'image':df.iloc[root_row + 5, root_column].values[0]
    }
    if data['audio']=='':
        data['audio'] = df.iloc[root_row + 4, root_column-1].values[0]
    if data['image']=='':
        data['image'] = df.iloc[root_row + 5, root_column-1].values[0]


    # correct file paths:
    data['image'] = data['image'].replace("'", '_')
    data['audio'] = data['audio'].replace("'", '_')

    return data





def get_time_frame(mp3_file):
    # convert mp3 to wav
    output = "aligner\\input\\temp_wav_file.wav"
    cmd = 'converter\\bin\\converter ' \
        + '-y -i "{}" -vn -ar 44100 -ac 1 -b:a 16k "{}"'.format(mp3_file, output)
    os.system(cmd)

    # create transcript file
    transcript = re.sub(r'(.+/)|(.+\\)','',mp3_file).replace('.mp3','')
    transcript = re.sub(r'\s+',' ',transcript)
    with open("aligner\\input\\temp_wav_file.txt", 'w') as file:
        file.write(transcript)

    # run aligner
    cmd = 'echo NO | aligner\\bin\\mfa_align ' \
         + 'aligner\\input\\ ' \
         + 'aligner\\librispeech-lexicon.txt ' \
         + 'aligner\\pretrained_models\\english.zip ' \
         + 'aligner\\output\\'
    os.system(cmd)

    # get output TextGrid
    with open('aligner\\output\\input\\temp_wav_file.TextGrid', 'r') as file:
        textgrid = file.read()
    
    # textgrid to list
    textgrid = re.sub('\t','',textgrid)
    textgrid = re.findall('name = "words".+name = "phones"',textgrid, flags=re.DOTALL)[0]
    intervals = re.findall(r'intervals \[.*?\].*?text = ".*?"', textgrid, flags=re.DOTALL)

    # convert textgrid to time frame
    time_mark = 0
    time_frame = {}
    words = transcript.split()
    word_count = 0
    for interval in intervals: 
        text = re.findall('".*"', interval)[0].replace('"','')
        if text!='':
            new_time_mark = re.findall('xmax = .*', interval)[0].replace('xmax = ','')
            time_length = float(new_time_mark) - time_mark
            time_mark = float(new_time_mark)
            time_frame[words[word_count].lower()] = str(round(time_length,2))
            word_count += 1
    print(time_frame)
    return time_frame



class TopKidDo(object):
    """docstring for TopKidDo"""
    def __init__(self, session):
        self.session = session
        self.login()

    def login(self):
        headers = {
            'Connection': 'keep-alive',
            'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.66 Safari/537.36',
            'Content-Type': 'application/json',
            'Accept': '*/*',
            'Origin': 'http://admin.topkiddovn.com',
            'Referer': 'http://admin.topkiddovn.com/',
            'Accept-Language': 'vi-VN,vi;q=0.9,fr-FR;q=0.8,fr;q=0.7,en-US;q=0.6,en;q=0.5',
        }
        data = '{"username":"admin@admin.com","password":"admin","remember":true}'
        url = 'http://backend.topkiddovn.com/users/login'
        while True:
            response = self.session.post(url, headers=headers, data=data, verify=False)
            if response.status_code==200: 
                print('>>>>>>>>>> logged in successfully')
                break
            else:
                print('>>>>>>>>>> failed to log in', response.status_code)
                time.sleep(1)

    def add_word(self, content, language, tags, position, animation=1, resources=''):
        ''' add new slide word
        > content: word
        > positon: 4 => Text invisible, 5 => image invisible
        > animation: 1 => flicker, 2 => fade in
        '''
        headers = {
            'Connection': 'keep-alive',
            'Authorization': 'Bearer eyJhbGciOiJIUzI1NiJ9.NjAxNmM3ZjdjYmM1MzkwM2JlZDI0MjZj.hwlBIiINCUSDQOWG69onlsgg-bbvNUXxLalhb_upGfo',
            'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.66 Safari/537.36',
            'Content-Type': 'application/json',
            'Accept': '*/*',
            'Origin': 'http://admin.topkiddovn.com',
            'Referer': 'http://admin.topkiddovn.com/administrator/contentmenu',
            'Accept-Language': 'en-US,en;q=0.9',
        }
        languages = {'English - American':2}
        positions = {'Text invisible':4, 'Image invisible':5}
        data = '''{{"data":{{"language":{language}, 
            "tags":"{tags}",
            "content":"{content}",
            "contentPosition":{position},
            "isShowFinger":false,
            "sizeContent":15,
            "colorContent":"#D0021B",
            "animationContent":{animation},
            "resources":[{resources}],
            "type":1}}}}'''.format(tags=tags, content=content, language=languages[language], position=positions[position], animation=animation, resources=resources)

        url = 'http://backend.topkiddovn.com/lessions/create_content_item'
        while True:
            response = self.session.post(url, headers=headers, data=data, verify=False)
            if response.status_code==200: 
                print('>>>>>>>>>> added word successfully:', content, tags)
                break
            else:
                print('>>>>>>>>>> failed to add word', content, tags , response.status_code)
                time.sleep(1)
                print(response.text)


    def upload_resource(self, filepath, tags='null', update=True):
        # correct file name
        filepath = filepath.replace("'",'_')
        
        # get file name from filepath (note: windows and linux have different splashes)
        filename = re.sub(r'(.+/)|(.+\\)','',filepath)
        
        # check if the file exists
        if not os.path.exists(filepath):
            print('>>>>>>>>>> file not found: ' + filepath + ', checking for closest filepath ...')
            dirpath = filepath.replace(filename,'')
            filenames = os.listdir(dirpath)
            for each_filename in filenames:
                if each_filename.lower().replace(' ','')==filename.lower().replace(' ',''):
                    filename = each_filename
                    filepath = dirpath + filename
                    break
            if not os.path.exists(filepath):
                raise Exception('>>>>>>>>>> file not found: ' + filepath + ', no alternatives')
            else:
                print('>>>>>>>>>> found an alternative: ' + filepath)

        # check if uploaded
        with open('upload_log.txt', 'r') as file:
            txt = file.read()
            if filename in txt:
                _id = re.findall(filename + '.+', txt)[0].replace(filename+':','')
                print('>>>>>>>>>> got from log:', filename)
                return _id

        boundary = '----WebKitFormBoundary' \
                   + ''.join(random.sample(string.ascii_letters + string.digits, 16))

        url = 'http://backend.topkiddovn.com/resources/upload_resource_local'
        if filename[-3:].lower()=='png':
            filetype = 'image/png'
        elif filename[-3:].lower()=='wma':
            filetype = 'audio/x-ms-wma'
        elif filename[-3:].lower()=='mp3':
            filetype = 'audio/mpeg'
        elif filename[-3:].lower()=='jpg':
            filetype = 'image/jpeg'
        elif filename[-3:].lower()=='mp4':
            filetype = 'video/mp4'
            
        mp_encoder = MultipartEncoder(
            fields={
                # plain file object, no filename or mime type produces a
                # Content-Disposition header with just the part name
                'file': (filename, open(filepath, 'rb'), filetype)}
        )

        headers = {
            'Connection': 'keep-alive',
            'Authorization': 'eyJhbGciOiJIUzI1NiJ9.NjAxNzUyNDZjYmM1MzkwM2JlZDI0Mjg2.mJ3qzE_NgsGJ_atPIw7zmEUfjYLsu4-ZuY2BO8XjEOg',
            'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.66 Safari/537.36',
            'Content-Type': mp_encoder.content_type,
            'Accept': '*/*',
            'Origin': 'http://admin.topkiddovn.com',
            'Referer': 'http://admin.topkiddovn.com/',
            'Accept-Language': 'vi-VN,vi;q=0.9,fr-FR;q=0.8,fr;q=0.7,en-US;q=0.6,en;q=0.5',
        }
        while True:
            response = self.session.post(
                url,
                data=mp_encoder,  # The MultipartEncoder is posted as data, don't use files=...!
                # The MultipartEncoder provides the content-type header with the boundary:
                headers=headers
            )
            if response.status_code==200: 
                print('>>>>>>>>>> uploaded successfully:', filename)
                break
            else:
                print('>>>>>>>>>> failed to upload:', filename, response.status_code)
                time.sleep(1)

        _id = re.findall('"_id":".+?"',str(response.text))[0].replace('"_id":"','').replace('"','')
        
        # update if the flag update is set (default)
        if update: self.update_resource(_id, filename, tags)
        
        # write the file name and id to a log file use next times
        with open('upload_log.txt', 'a') as file:
            file.write('\nupload:' + filename + ':' + _id + '\n')
        return _id


    def update_resource(self, _id, filename, tags='null'):
        headers = {
            'Connection': 'keep-alive',
            'Authorization': 'Bearer eyJhbGciOiJIUzI1NiJ9.NjAxNzUyNDZjYmM1MzkwM2JlZDI0Mjg2.mJ3qzE_NgsGJ_atPIw7zmEUfjYLsu4-ZuY2BO8XjEOg',
            'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.66 Safari/537.36',
            'Content-Type': 'application/json',
            'Accept': '*/*',
            'Origin': 'http://admin.topkiddovn.com',
            'Referer': 'http://admin.topkiddovn.com/',
            'Accept-Language': 'vi-VN,vi;q=0.9,fr-FR;q=0.8,fr;q=0.7,en-US;q=0.6,en;q=0.5',
        }

        #data = '{"resource_id":"{_id}","updateData":{"name":{filename},"tags":[null]}}'.format(_id=_id,filename=filename)
        # convert tags in str or list to proper format for data
        if type(tags)==str:
            if tags!='null':
                tags = '"{}"'.format(tags)
        else:
            tags = '"{}"'.format('","'.join(tags))

        data = '{{"resource_id":"{}","updateData":{{"name":"{}","tags":[{}]}}}}'.format(_id, filename, tags)
        while True:
            response = self.session.post('http://backend.topkiddovn.com/resources/update_resource', headers=headers, data=data, verify=False)
            if response.status_code==200: 
                print('>>>>>>>>>> updated resources successfully: ', filename, _id)
                break
            else:
                print('>>>>>>>>>> failed to update resource', filename, _id, response.status_code)
                time.sleep(1)


    def create_normal_slides(self, df_word, word):
        data = get_word_data(df_word, word)
        # if can not find word
        if data==[]:
            raise Exception('>>>>>>>>>> can not find data for the word: ' + word)

        for step_data in data:

            content = step_data['content']
            position = step_data['position']
            language = step_data['language']
            tags = step_data['tags']
            audio = step_data['audio']
            image = step_data['image']
            video = step_data['video']

            # check if upload this data
            with open('upload_log.txt', 'r') as file:
                if str(step_data) in file.read():
                    print('>>>>>>>>>> this content "{}" (tags: {}) has already been created before, please check'.format(content,tags))
                    continue

            ids = []
            for filename in (audio, image, video):
                filepath = 'TOPKIDDO\\Words\\' + filename
                if filename!='':
                    ids.append('"' + self.upload_resource(filepath) + '"')
            resources = ','.join(ids)
            print(resources)

            self.add_word(content=content, 
                position=position, 
                language=language,
                tags=tags,
                resources=resources)

            # log step_data to a file to avoid duplicates
            with open('upload_log.txt', 'a') as file:
                file.write('\nadd word:' + str(step_data) + '\n')


    def create_letter_resource(self, tags):
        headers = {
            'Connection': 'keep-alive',
            'Authorization': 'Bearer eyJhbGciOiJIUzI1NiJ9.NjAxNzUyNDZjYmM1MzkwM2JlZDI0Mjg2.mJ3qzE_NgsGJ_atPIw7zmEUfjYLsu4-ZuY2BO8XjEOg',
            'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.66 Safari/537.36',
            'Content-Type': 'application/json',
            'Accept': '*/*',
            'Origin': 'http://admin.topkiddovn.com',
            'Referer': 'http://admin.topkiddovn.com/',
            'Accept-Language': 'vi-VN,vi;q=0.9,fr-FR;q=0.8,fr;q=0.7,en-US;q=0.6,en;q=0.5',
        }
        data = '{{"data":{{"letter":"{}"}}}}'.format(tags)
        response = self.session.post('http://backend.topkiddovn.com/lessions/create_letter_resource', headers=headers, data=data, verify=False)
        tag_id = re.findall('"_id":".+?"',str(response.text))[0].replace('"_id":"','').replace('"','')
        return tag_id

    def add_resource_to_letter(self, resource_id, tag_id):
        headers = {
            'Connection': 'keep-alive',
            'Authorization': 'Bearer eyJhbGciOiJIUzI1NiJ9.NjAxNzUyNDZjYmM1MzkwM2JlZDI0Mjg2.mJ3qzE_NgsGJ_atPIw7zmEUfjYLsu4-ZuY2BO8XjEOg',
            'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.66 Safari/537.36',
            'Content-Type': 'application/json',
            'Accept': '*/*',
            'Origin': 'http://admin.topkiddovn.com',
            'Referer': 'http://admin.topkiddovn.com/',
            'Accept-Language': 'vi-VN,vi;q=0.9,fr-FR;q=0.8,fr;q=0.7,en-US;q=0.6,en;q=0.5',
        }
        data = '{{"data":{{"resourceId":"{}","letterResourceId":"{}"}}}}'.format(resource_id, tag_id)
        response = self.session.post('http://backend.topkiddovn.com/lessions/add_resource_to_letter', headers=headers, data=data, verify=False)


    def add_letter_resource_to_content(self, content_id, tag_id):
        headers = {
            'Connection': 'keep-alive',
            'Authorization': 'Bearer eyJhbGciOiJIUzI1NiJ9.NjAxNzUyNDZjYmM1MzkwM2JlZDI0Mjg2.mJ3qzE_NgsGJ_atPIw7zmEUfjYLsu4-ZuY2BO8XjEOg',
            'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.66 Safari/537.36',
            'Content-Type': 'application/json',
            'Accept': '*/*',
            'Origin': 'http://admin.topkiddovn.com',
            'Referer': 'http://admin.topkiddovn.com/',
            'Accept-Language': 'vi-VN,vi;q=0.9,fr-FR;q=0.8,fr;q=0.7,en-US;q=0.6,en;q=0.5',
        }
        data = '{{"data":{{"contentItemId":"{}","letterResourceId":"{}"}}}}'.format(content_id, tag_id)
        response = self.session.post('http://backend.topkiddovn.com/lessions/add_letter_resource_to_content', headers=headers, data=data, verify=False)


    def add_multi(self, tags, resources='', language='English - American'):
        headers = {
            'Connection': 'keep-alive',
            'Authorization': 'Bearer eyJhbGciOiJIUzI1NiJ9.NjAxNzUyNDZjYmM1MzkwM2JlZDI0Mjg2.mJ3qzE_NgsGJ_atPIw7zmEUfjYLsu4-ZuY2BO8XjEOg',
            'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.66 Safari/537.36',
            'Content-Type': 'application/json',
            'Accept': '*/*',
            'Origin': 'http://admin.topkiddovn. com',
            'Referer': 'http://admin.topkiddovn.com/',
            'Accept-Language': 'vi-VN,vi;q=0.9,fr-FR;q=0.8,fr;q=0.7,en-US;q=0.6,en;q=0.5',
        }
        languages = {'English - American':2}
        data = '''{{"data":{{"language":{language}, 
            "tags":"{tags}",
            "type":3,
            "resources":[{resources}]}}}}'''.format(tags=tags, language=languages[language], resources=resources)
        while True:
            response = self.session.post('http://backend.topkiddovn.com/lessions/create_content_item', headers=headers, data=data, verify=False)
            if response.status_code==200: 
                print('>>>>>>>>>> added multi successfully:', tags)
                break
            else:
                print('>>>>>>>>>> failed to add multi', tags , response.status_code)
                time.sleep(1)
                print(response.text)
        content_id = re.findall('"_id":".+?"',str(response.text))[0].replace('"_id":"','').replace('"','')
        return content_id

    def create_special_slide(self, df_word, word1, word2):
        data = {
            'word1':row['word_1'],
            'word2':row['word_2']
        }

        # check if upload this data
        tags = data['word1'] + ' - ' + data['word2']
        with open('upload_log.txt', 'r') as file:
            if str(data) in file.read():
                print('>>>>>>>>>> this content "{tags}" (tags: {tags}) has already been created before, please check'.format(tags=tags))
                return

        dir_path = 'TOPKIDDO\\Words\\'
        content_id = self.add_multi(tags=tags)

        for word in (data['word1'],data['word2']):
            word_data = get_word_data(df_word, word)
            # if can not find word
            if word_data==[]:
                raise Exception('>>>>>>>>>> can not find data for the word: ' + word)

            audio = word_data[0]['audio']
            image = word_data[0]['image']
            for filename in (image, audio):
                tag_id = self.create_letter_resource(word)
                resource_id = self.upload_resource(dir_path + filename)
                self.add_resource_to_letter(resource_id, tag_id)
                self.add_letter_resource_to_content(content_id, tag_id)


        # log step_data to a file to avoid duplicates
        with open('upload_log.txt', 'a') as file:
            file.write('\nadd multi:' + str(data) + '\n')
        print('>>>>>>>>>> successfully created multi slide: ', tags)


    def add_sentence(self, content, language, tags, position, animation=1, resources=''):
        ''' add new slide word
        > content: word
        > positon: 4 => Text invisible, 5 => image invisible
        > animation: 1 => flicker, 2 => fade in
        '''
        headers = {
            'Connection': 'keep-alive',
            'Authorization': 'Bearer eyJhbGciOiJIUzI1NiJ9.NjAyMjBiOTdjYmM1MzkwM2JlZDI0M2Vi.-_ujV_Hheeuvnx2z2iR80xRMiNHSOJpID7Uspc_uI00',
            'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.146 Safari/537.36',
            'Content-Type': 'application/json',
            'Accept': '*/*',
            'Origin': 'http://admin.topkiddovn.com',
            'Referer': 'http://admin.topkiddovn.com/',
            'Accept-Language': 'vi-VN,vi;q=0.9,fr-FR;q=0.8,fr;q=0.7,en-US;q=0.6,en;q=0.5',
        }
        languages = {'English - American':2}
        positions = {'Text invisible':4, 'Image invisible':5}
        resources = '"' + '","'.join(resources) + '"'
        data = '''{{"data":{{"language":{language}, 
            "tags":"{tags}",
            "content":"{content}",
            "contentPosition":{position},
            "isShowFinger":false,
            "sizeContent":15,
            "colorContent":"#D0021B",
            "highlightColor":"#F6E40A",
            "animationContent":{animation},
            "resources":[{resources}],
            "type":2,
            "highlight":[]}}}}'''.format(tags=tags, content=content, language=languages[language], position=positions[position], animation=animation, resources=resources)

        url = 'http://backend.topkiddovn.com/lessions/create_content_item'
        while True:
            response = self.session.post(url, headers=headers, data=data, verify=False)
            if response.status_code==200: 
                print('>>>>>>>>>> added sentence successfully:', content, tags)
                break
            else:
                print('>>>>>>>>>> failed to add sentence', content, tags , response.status_code)
                time.sleep(1)
                print(response.text)
        content_id = re.findall('"_id":".+?"',str(response.text))[0].replace('"_id":"','').replace('"','')
        return content_id

    def add_time_frame(self, content_id, data, audio_path):
        time_frame = get_time_frame(audio_path)
        headers = {
            'Connection': 'keep-alive',
            'Authorization': 'Bearer eyJhbGciOiJIUzI1NiJ9.NjAyMjBiOTdjYmM1MzkwM2JlZDI0M2Vi.-_ujV_Hheeuvnx2z2iR80xRMiNHSOJpID7Uspc_uI00',
            'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.146 Safari/537.36',
            'Content-Type': 'application/json',
            'Accept': '*/*',
            'Origin': 'http://admin.topkiddovn.com',
            'Referer': 'http://admin.topkiddovn.com/',
            'Accept-Language': 'vi-VN,vi;q=0.9,fr-FR;q=0.8,fr;q=0.7,en-US;q=0.6,en;q=0.5',
        }
        time = '['
        for word in data['content'].split():
            time += '{{\\"key\\":\\"{}\\",\\"time\\":{}}},'.format(word, time_frame[word.lower().replace("'",'_')])
        time = time[:-1] + ']'


        data = '''{{"contentItemId":"{}","timeFrame":"{}"}}'''.format(content_id, time)
        response = self.session.post('http://backend.topkiddovn.com/lessions/edit_content_item', 
            headers=headers, data=data, verify=False)

    def create_sentence_slide(self, df_sentence, df_word, sentence):
        data = get_sentence_data(df_sentence, sentence)
        # if can not find sentence
        if data==[]:
            raise Exception('>>>>>>>>>> can not find data for the sentence: ' + sentence)
 
        # check if upload this data
        with open('upload_log.txt', 'r') as file:
            if str(data) in file.read():
                print('>>>>>>>>>> this content "{content}" (tags: {tags}) has already been created before, please check'.format(content=data['content'],tags=data['tags']))
                return

        dir_path = 'TOPKIDDO\\Sentences\\'

        filename = data['audio']
        audio_sentence_id = self.upload_resource(dir_path + filename, tags=filename[:-4])
        filename = data['image']
        image_sentene_id = self.upload_resource(dir_path + filename, tags=filename[:-4])


        content = data['content']
        position = data['position']
        language = data['language']
        tags = data['tags']

        content_id = self.add_sentence(content, language, tags, position, resources=(audio_sentence_id,image_sentene_id))
        for word in content.split():
            tag_id = self.create_letter_resource(word)
            word_data = get_word_data(df_word, word)
            if word_data!=[]:
                audio = word_data[0]['audio']
                image = word_data[0]['image']
                for filename in (audio, image):
                    resource_id = self.upload_resource('TOPKIDDO\\Words\\' + filename)
                    self.add_resource_to_letter(resource_id, tag_id)
            self.add_letter_resource_to_content(content_id, tag_id)
        
        audio_path = dir_path + data['audio']
        self.add_time_frame(content_id, data, audio_path)
        # log step_data to a file to avoid duplicates
        with open('upload_log.txt', 'a') as file:
            file.write('\nadd sentence:' + str(data) + '\n')
        print('>>>>>>>>>> successfully created sentence slide: ', content, '(tags: {})'.format(tags))










#===========================================
df_word = input_excel_database('Add words')
df_task = input_excel_database('Tasks')
df_sentence = input_excel_database('Add sentences')

with requests.Session() as session:
    topkiddo = TopKidDo(session)
    for index, row in df_task.iterrows():
        if row['task']=='word':
            word = row['word_1']
            topkiddo.create_normal_slides(df_word, word)

        elif row['task']=='multi':
            word1 = row['word_1']
            word2 = row['word_2']
            topkiddo.create_special_slide(df_word, word1, word2)

        elif row['task']=='sentence':
            sentence = row['sentence']
            topkiddo.create_sentence_slide(df_sentence, df_word, sentence)



