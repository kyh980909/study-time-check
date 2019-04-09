import requests
from bs4 import BeautifulSoup
from requests import get  # to make GET request
from openpyxl import load_workbook
import os
import time
from time import strftime


def download(url, filename):
    with open(filename, "wb") as file:  # open in binary mode
        response = get(url)  # get request
        file.write(response.content)  # write to file


def file_check(filename):   # 파일 이름을 매개변수로 보내서 파일이 있으면 True 없으면 False 반환
    return os.path.exists(filename)


start_time = time.time()

login_url = 'http://e-portfolio.bible.ac.kr/Templates/sFTLogin.aspx'

user = '201704005'  # 아이디
pw = '2017040'  # 비밀번호

session = requests.session()
cookies = {'MEM_ID': user}  # 쿠키 생성

params = dict()
params['txtOprID'] = user
params['txtOprPW'] = pw

res = session.post(login_url, data=params)

res.raise_for_status()

study_url = 'http://e-portfolio.bible.ac.kr/blog/blogMain.aspx?blogid=study-ctl&catseqno=21369'
res = session.get(study_url, cookies=cookies)  # 쿠키 넣기

res.raise_for_status()

soup = BeautifulSoup(res.text, 'html.parser')

excel = soup.select_one('#ctl00_cphMain_dlArtList_ctl00_lbAttFile > a')

file_name_tag = str(excel.find('font'))
file_name = file_name_tag[21:60]  # 파일 이름 슬라이싱

if file_check(file_name):
    print('파일이 있어서 다운받지 않습니다.')
    pass
else:
    if str(excel).find('출석현황') != -1:  # 출석현황 파일인지 아닌지 구분
        down_url = 'http://e-portfolio.bible.ac.kr' + excel.get('href')  # 태그안 href속성 값 가져오기

        download(down_url, file_name)

        print(down_url)
        print('파일 다운 완료')
    else:
        print('출석현황 파일이 아닙니다.')

# 여기까지 출석현황 파일 다운로드

# 여기부터 엑셀 읽기 시작

load_wb = load_workbook(file_name, data_only=True)
load_ws = load_wb['Sheet1']

print(load_ws['A1'].value)

print('걸린시간:', time.time() - start_time)