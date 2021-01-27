# Programmed by Chen Lisi

from urllib.request import urlopen
from bs4 import BeautifulSoup

import xlrd
import xlwt
from xlutils.copy import copy

import time

import wikipedia

# VPN requirement:
# 1. 必须使用全局代理
# 2. 服务器可能有反扒， 需准备多个IP

# Born Date Extraction Rule:
# 1. Extract the text in tag <span>, whose attribute is bday
# 2. Remove Jr.
# 3. If length of name > 3, remove the second word (middle name)

# Education Extraction:
# 1. Extract the text in tag <td><a> ... </a></td>
# 2. The text has word 'University' or 'College'

# Link Extraction:
# 1. Just return the search url link of wikipedia

# Home state Extraction:
# 1. Extract the first state that appears in html, while falling into <a>...</a>

# Summary Extraction:
# Utilize the package name called 'wikipedia' and directly request summary
# Does not work for multi-meaning enquiry


def parse_data(raw_name):
    # URL generation
    name = raw_name.split()
    # remove Jr
    for j in range(0, len(name)):
        if name[j] == 'Jr':
            del name[j]
    # remove middle name
    if len(name) > 2:
        del name[1]
    name = " ".join(name)
    name = name.replace(" ", "_")
    url = 'https://en.wikipedia.org/wiki/'+ name
    print(url)
    html = urlopen(url)
    soup = BeautifulSoup(html, 'html.parser')
    # Search born data
    s1 = soup.find_all('span', attrs={"class": "bday"})  # 查找span class为bday的字符串
    result = [span.get_text() for span in s1]
    print(result)
    if len(result) > 1:
        born_result = result[0]
    born_result = result
    # Search education
    edu_list = []
    s2 = soup.find_all('td')  # 查找td的字符串
    for s3 in s2:
        for result in s3.find_all('a'):
            school = result.get_text()
            if school.find('University') != -1:
                edu_list.append(school)
            if school.find('College') != -1:
                edu_list.append(school)
    # Search home state
    state_list = []
    for result in soup.find_all('a'):
        state = result.get_text()
        for state_name in {'Alabama', 'Alaska', 'Arizona', 'Arkansas',
                           'California', 'Colorado', 'Connecticut', 'Delaware',
                           'Florida', 'Georgia', 'Hawaii', 'Idaho', 'Illinois',
                           'Indiana', 'Iowa', 'Kansas', 'Kentucky', 'Louisiana',
                           'Maine', 'Maryland', 'Massachusetts', 'Michigan',
                           'Minnesota', 'Mississippi', 'Missouri', 'Montana',
                           'Nebraska', 'Nevada', 'New Hampshire', 'New Mexico',
                           'North Carolina', 'North Dakota', 'Ohio', 'Oklahoma',
                           'Oregon', 'Pennsylvania', 'Rhode Island', 'South Carolina',
                           'South Dakota', 'Tennessee', 'Texas', 'Utah', 'Vermont',
                           'Virginia', 'Washington', 'West Virginia', 'Wisconsin', 'Wyoming',
                           'Virginia', 'New York', 'New Jersey'}:
            if state.find(state_name) != -1:
                state_list.append(state)
    if len(state_list) > 1:
        state_result = state_list[0]
    else:
        state_result = state_list
    print(state_result)
    summary = []
    try:
        summary = wikipedia.summary(raw_name)
    except Exception as e:
        pass
    print(summary)
    return born_result, edu_list, url, state_result, summary


def process(start_index, stop_index):
    workbook = xlrd.open_workbook("politicianlist_2015.xlsx")  # 文件路径
    worksheet = workbook.sheet_by_name("politicianlist_2015")

    filename = 'output.xls'  # 文件名
    rb = xlrd.open_workbook(filename, formatting_info=True)
    book = copy(rb)
    # book = xlwt.Workbook(encoding="utf-8", style_compression=0)
    sheet = book.get_sheet(0)
    for i in range(start_index, stop_index):
        name = worksheet.cell(i, 2).value
        print('i: ', i, 'name: ', name)
        # Parse Data
        born_date, edu_list, url, state_result, summary = parse_data(name)
        # write born date
        sheet.write(i, 0, born_date)
        # write edu list
        if len(edu_list) > 3:
            edu_list = edu_list[0:3]
        edu_list = list(set(edu_list))
        print(edu_list)
        for j in range(0, len(edu_list)):
            sheet.write(i, 4 + j, edu_list[j])
        # write link
        sheet.write(i, 8, url)
        # write home state
        sheet.write(i, 3, state_result)
        # write summary
        sheet.write(i, 7, summary)
        # Save every ...
        if i % 10 == 0:
            book.save('output.xls')
            print('save: ', i)
        if i == (stop_index - 1):
            book.save('output.xls')
            print('save: ', i)
            return
        time.sleep(2)
    return


process(10, 319)
