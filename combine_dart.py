# -*- coding: utf-8 -*-
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from PyQt5 import uic
from PyQt5 import QtCore
import urllib
import zipfile
from bs4 import BeautifulSoup
import os
import json
import pandas as pd
import sys

form_class = uic.loadUiType("dart_api.ui")[0]

class combine_dart(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        self.key = "172fb1624826c2efe673916cc1de96a9eab26961"  #oli@radexel.com
        self.key = "db997885049ef0ea762ad6098e3791d0b6e7c75e"  #henryjeong84268@gmail.com
        # self.key = "604ebad865e4d98c988ff43d37e12c79c887c138"  #pop84268@hotmail.com
        # self.key = "eaf7f1ed3cb845e57fead9cc73b3984698ca36d4"  #hannhyun@gmail.com
        # self.key = "bcafb129d6606e076425808121d7c9a07f5ecabc"  #pop84268@korea.ac.kr
        # self.key = "09e2f1db7f37f6d44f8c6f2b984da99e96c1d4aa"  #jjjang2001@naver.com
        
        # self.key = "79079cd2f961e086e49a006ddfe6244b1413ed2b "  # perennial247@gmail.com 
        # self.key = "395f336f01f7f44f495306cf30399ec8eedbfe40"  # rhgustjr4645@naver.com 
        # self.key = "106abda00148e39079fe91101390f1ac421c2147"  # perennial.247@daum.net 
        # self.key = "6a3fec154e270af7270f63d9a4d140dbd7b1b7a6"  # gomonggu@gmail.com 

        url1 = "https://opendart.fss.or.kr/api/corpCode.xml?crtfc_key={}".format(self.key)
        urllib.request.urlretrieve(url1, 'firmlist.zip')

        with zipfile.ZipFile('firmlist.zip', 'r') as zip_ref:
            zip_ref.extractall()
        fm = open('CORPCODE.xml', encoding="utf-8").read()
        fm = str(str(fm).replace("        ","").replace("    ","").replace("\n","").replace("<list><corp_code>",",").replace("</corp_code><corp_name>",",").replace("</corp_name><stock_code>",",").replace("</stock_code><modify_date>",",").replace("</modify_date>","").replace("</result>","").replace("<result>",""))
        self.fm2 = fm.split("</list>,")
        print(len(self.fm2))
        # print(fm)
        # # print(QTime.currentTime())
        # print("BeautifulSoup 실행")
        #
        # self.B = BeautifulSoup(fm, 'lxml')

        self.firm_items = {}
        print("DART API 실행")
        self.enter_key.clicked.connect(self.make_file_one)
        self.enter_key_2.clicked.connect(self.arrange_date_one)

        self.enter_key_3.clicked.connect(self.make_file_all)
        self.enter_key_4.clicked.connect(self.arrange_date_all)


    def make_file_all(self):
        current_yr = int(QDateTime.currentDateTime().toString("yyyy"))
        n = 0
        for i in reversed(self.fm2):
            i = i.split(",")

            corp_code = i[0]
            corp_name = i[1]
            stock_code = i[2]
            mod_date = i[3]

            try :
                print("상장여부", int(stock_code))
                n = n+1
                #
                # corp_name = str(self.lineEdit_3.text())
                # print(corp_name, "파일 생성 중")
                for i in range(current_yr,current_yr-2,-1) :
                    yr = str(i)
                    for j in ["q1","half year","q3",'year'] :
                        bungi = j
                        for k in ["현금흐름표","재무상태표","포괄손익계산서"] :
                            bogoseo = k
                            # print(corp_name, yr, bungi, "c", bogoseo)
                            try :
                                self.get_periodic_financial_statements(corp_name, yr, bungi, "c", bogoseo)

                            except :
                                pass
                                # print("보고서 없음 : ",corp_name, yr, bungi, "c", bogoseo)


                print(n, corp_name, "파일 생성 완료")
            except :
                pass
                # print("비상장",corp_name)
        print(n)

    def arrange_date_all(self):
        current_yr = int(QDateTime.currentDateTime().toString("yyyy"))
        n = 0

        directory = "dart\\"
        filelist = os.listdir(directory)
        print(filelist)

        for i in filelist :

            # corp_name = i[1]
            stock_code = i

            try :
                n = n+1

                fr = []
                for kind in ['cash_flow','balance_sheet','income_statement'] :
                    directory2 = 'dart\\' + str(stock_code) + '\\financial_statements\\' + kind + "\\"
                    filelist = os.listdir(directory2)

                    for i in filelist:
                        fiscal_yr = i[0:i.find("_")]
                        try:
                            int(fiscal_yr)
                            fr.append(int(fiscal_yr))
                        except:
                            pass
                            # print(i)
                            # os.remove(directory2+i)
                fr = list(set(fr))
                fr.sort()
                print(fr)
                for yr in fr:
                    yr = str(yr)

                    for k in ["현금흐름표","재무상태표","포괄손익계산서"] :
                            if k == "현금흐름표" :
                                kind1 = 'cash_flow'
                                kind2 = "_c_cash_flow.csv"
                            elif k == "재무상태표" :
                                kind1 = 'balance_sheet'
                                kind2 = "_c_balance_sheet.csv"
                            else :
                                kind1 = 'income_statement'
                                kind2 = "_c_income_statement.csv"

                            a = []

                            for j in ["q1", "half year", "q3", 'year'] :

                                directory2 = 'dart\\' + str(stock_code) + '\\financial_statements\\' + kind1 + "\\"
                                filelist = os.listdir(directory2)

                                for file in filelist :
                                    # print(yr+"_"+j+kind2[:-4])
                                    if yr+"_"+j+kind2[:-4] in file :   #92_half year_c_balance_sheet_20200814
                                        try :
                                            bungi = j
                                            df2 = pd.read_csv('dart\\'+stock_code+'\\financial_statements\\'+kind1+"\\"+file,engine='python',index_col=0)
                                            print('dart\\'+stock_code+'\\financial_statements\\'+kind1+"\\"+file)

                                            if kind1 == 'income_statement':
                                                if bungi == "q1" :
                                                    df = df2
                                                    df['q1'] = df['금액']
                                                    a.append("q1")
                                                elif bungi == "half year" :
                                                    df['q2'] = df2["금액"]
                                                    a.append("q2")
                                                elif bungi == "q3" :
                                                    df['q3'] = df2["금액"]
                                                    a.append("q3")
                                                else :
                                                    df['total'] = df2["금액"]
                                                    df['q4'] = df2["금액"] - df['q3'] - df['q2'] - df['q1']
                                                    a.append("q4")
                                                    a.append("total")
                                            elif kind1 == 'cash_flow' :
                                                if bungi == "q1" :
                                                    df = df2
                                                    df['q1'] = df['금액']
                                                    a.append("q1")
                                                elif bungi == "half year" :
                                                    df['q2'] = df2["금액"] - df['q1']
                                                    a.append("q2")
                                                elif bungi == "q3" :
                                                    df['q3'] = df2["금액"] - df['q2'] - df['q1']
                                                    a.append("q3")
                                                else :
                                                    df['total'] = df2["금액"]
                                                    df['q4'] = df2["금액"] - df['q3'] - df['q2'] - df['q1']
                                                    a.append("q4")
                                                    a.append("total")
                                            else :
                                                if bungi == "q1" :
                                                    df = df2
                                                    df['q1'] = df['금액']
                                                    a.append("q1")
                                                elif bungi == "half year" :
                                                    df['q2'] = df2["금액"]
                                                    a.append("q2")
                                                elif bungi == "q3" :
                                                    df['q3'] = df2["금액"]
                                                    a.append("q3")
                                                else :
                                                    df['total'] = df2["금액"]
                                                    df['q4'] = df2["금액"]
                                                    a.append("q4")
                                                    a.append("total")

                                        except :
                                            print("보고서 없음",yr+"_"+bungi+kind2)
                            df = df[a]
                            df.to_csv('dart\\'+stock_code+'\\financial_statements\\'+kind1+"\\"+kind1+"_"+yr+".csv", encoding='euc-kr')
            except :
                pass
                # print("비상장",corp_name)
        print(n)

    def arrange_date_one(self):
        current_yr = int(QDateTime.currentDateTime().toString("yyyy"))


        corp_name = str(self.lineEdit_3.text())
        corp_code = self.get_corp_code(corp_name)  # 기업 코드 가져오기 ---> API 입력용
        stock_code = self.get_stock_code(corp_name)  # 종목 코드 가져오기 ---> 파일 저장 이름 용

        #cash_flow
        for i in range(current_yr,current_yr-2,-1) :
            yr = str(i)

            for k in ["현금흐름표","재무상태표","포괄손익계산서"] :
                if k == "현금흐름표" :
                    kind1 = 'cash_flow'
                    kind2 = "_c_cash_flow.csv"
                elif k == "재무상태표" :
                    kind1 = 'balance_sheet'
                    kind2 = "_c_balance_sheet.csv"
                else :
                    kind1 = 'income_statement'
                    kind2 = "_c_income_statement.csv"
                print(kind1+"\\"+yr+"_q1"+kind2)

                a = []
                for j in ["q1", "half year", "q3", 'year'] :
                    try :
                        bungi = j
                        df2 = pd.read_csv('dart\\'+stock_code+'\\financial_statements\\'+kind1+"\\"+yr+"_"+bungi+kind2,engine='python',index_col=0)
                        if kind1 == 'income_statement' :
                            if bungi == "q1" :
                                df = df2
                                df['q1'] = df['금액']
                                a.append("q1")
                            elif bungi == "half year" :
                                df['q2'] = df2["금액"]
                                a.append("q2")
                            elif bungi == "q3" :
                                df['q3'] = df2["금액"]
                                a.append("q3")
                            else :
                                df['total'] = df2["금액"]
                                df['q4'] = df2["금액"] - df['q3'] - df['q2'] - df['q1']
                                a.append("q4")
                                a.append("total")
                        elif kind1 == 'cash_flow' :
                            if bungi == "q1" :
                                df = df2
                                df['q1'] = df['금액']
                                a.append("q1")
                            elif bungi == "half year" :
                                df['q2'] = df2["금액"] - df['q1']
                                a.append("q2")
                            elif bungi == "q3" :
                                df['q3'] = df2["금액"]
                                a.append("q3")
                            else :
                                df['total'] = df2["금액"]
                                df['q4'] = df2["금액"] - df['q3'] - df['q2'] - df['q1']
                                a.append("q4")
                                a.append("total")
                        else :
                            if bungi == "q1" :
                                df = df2
                                df['q1'] = df['금액']
                                a.append("q1")
                            elif bungi == "half year" :
                                df['q2'] = df2["금액"]
                                a.append("q2")
                            elif bungi == "q3" :
                                df['q3'] = df2["금액"]
                                a.append("q3")
                            else :
                                df['total'] = df2["금액"]
                                df['q4'] = df2["금액"]
                                a.append("q4")
                                a.append("total")

                    except :
                        print("보고서 없음",yr+"_"+bungi+kind2)
                df = df[a]
                df.to_csv('dart\\'+stock_code+'\\financial_statements\\'+kind1+"\\"+kind1+"_"+yr+".csv", encoding='euc-kr')

    def make_file_one(self):
        current_yr = int(QDateTime.currentDateTime().toString("yyyy"))

        corp_name = str(self.lineEdit_3.text())
        print(corp_name, "파일 생성 중")
        for i in range(current_yr-1,current_yr-2,-1) :
            yr = str(i)
            for j in ["q1",'year',"half year","q3"] :
                bungi = j
                for k in ["현금흐름표","재무상태표","포괄손익계산서"] :
                    bogoseo = k
                    print(corp_name, yr, bungi, "c", bogoseo)
                    # try :
                    self.get_periodic_financial_statements(corp_name, yr, bungi, "c", bogoseo)
                    # except :
                    #     print("보고서 없음 : ",corp_name, yr, bungi, "c", bogoseo)
                # self.get_periodic_financial_statements(corp_name, "2021", "q1", "c", "재무상태표")
                # self.get_periodic_financial_statements(corp_name, "2021", "q1", "c", "포괄손익계산서")
        #
        # if target == "재무상태표":  # 파일 명을 영어로 해야 하기에....
        # elif target == "현금흐름표":
        # elif target == "포괄손익계산서":

        #
        # if reprt_type == "year":  # 사업보고서
        # elif reprt_type == "half year":  # 반기보고서
        # elif reprt_type == "q1":  # 1분기보고서
        # elif reprt_type == "q3":  # 3분기보고서

        print(corp_name, "파일 생성 완료")

    def get_firm_name(self, input_code):
        corp_code = self.B.find_all("corp_code")
        stock_code = self.B.find_all("stock_code")
        firm_name = self.B.find_all("corp_name")
        for C, S, F in zip(corp_code, stock_code, firm_name):
            # stock = S.get_text().strip()
            fcode = C.get_text().strip()
            name = F.get_text().strip()
            if fcode == input_code:
                return name

    def get_corp_code(self,input_firm):  # 반드시 일치하는 이름 입력해야 함.
        # corp_code = self.B.find_all("corp_code")
        # stock_code = self.B.find_all("stock_code")
        # firm_name = self.B.find_all("corp_name")
        for i in self.fm2 :
            i = i.split(",")
            corp_code = i[0]
            firm_name = i[1]
            stock_code = i[2]
            mod_date = i[3]
            # print(firm_name)
            if firm_name == input_firm:
                return corp_code


        # for C, S, F in zip(corp_code, stock_code, firm_name):
        #     # stock = S.get_text().strip()
        #     fcode = C.get_text().strip()
        #     name = F.get_text().strip()
        #     if name == input_firm:
        #         return fcode


    def get_corp_code_for_oli(self,stock_code):  # 반드시 일치하는 이름 입력해야 함.
        # corp_code = self.B.find_all("corp_code")
        # stock_code = self.B.find_all("stock_code")
        # firm_name = self.B.find_all("corp_name")
        for i in self.fm2 :
            i = i.split(",")
            corp_code = i[0]
            firm_name = i[1]
            stock_code_chk = i[2]
            mod_date = i[3]
            # print(firm_name)
            if stock_code == stock_code_chk:
                return corp_code


        # for C, S, F in zip(corp_code, stock_code, firm_name):
        #     # stock = S.get_text().strip()
        #     fcode = C.get_text().strip()
        #     name = F.get_text().strip()
        #     if name == input_firm:
        #         return fcode


    def get_stock_code(self,input_firm):
        # corp_code = self.B.find_all("corp_code")
        # stock_code = self.B.find_all("stock_code")
        # firm_name = self.B.find_all("corp_name")

        for i in self.fm2:
            i = i.split(",")

            corp_code = i[0]
            firm_name = i[1]
            stock_code = i[2]
            mod_date = i[3]

            if firm_name == input_firm:
                return stock_code




        # for C, S, F in zip(corp_code, stock_code, firm_name):
        #     stock = S.get_text().strip()
        #     # fcode = C.get_text().strip()
        #     name = F.get_text().strip()
        #     if name == input_firm:
        #         return stock

    # In[11]:

    # TEST CASE
    # print(type(get_stock_code("삼성전자")))

    # In[4]:

    def createDirectory(self,directory):
        try:
            if not os.path.exists(directory):
                os.makedirs(directory)
        except OSError:
            print("Error: Failed to create the directory.")

    # In[5]:

    def create_sub_Directory(self,sub_directory, path):
        try:
            if not os.path.exists(sub_directory):
                os.makedirs(sub_directory)
        except OSError:
            print("Error: Failed to create the directory.")

    # In[13]:

    # 정기공시 구하고 저장

    def get_periodic_financial_statements(self,firm_name, year, reprt_type, consolidated, target):
        if reprt_type == "year":  # 사업보고서
            code = "11011"
        elif reprt_type == "half year":  # 반기보고서
            code = "11012"
        elif reprt_type == "q1":  # 1분기보고서
            code = "11013"
        elif reprt_type == "q3":  # 3분기보고서
            code = "11014"

        stock_code = self.get_stock_code(firm_name)
        corp_code = self.get_corp_code_for_oli(stock_code) # 기업 코드 가져오기 ---> API 입력용
          # 종목 코드 가져오기 ---> 파일 저장 이름 용

        folder = 'dart'
        if consolidated == "c":
            fs_div = "CFS"  # 연결 재무제표

        elif consolidated == "o":
            fs_div = "OFS"  # 단일 재무제표

        url = "https://opendart.fss.or.kr/api/fnlttSinglAcntAll.json?crtfc_key={}&corp_code={}&bsns_year={}&reprt_code={}&fs_div={}".format(
            self.key, corp_code, year, code, fs_div)

        a = stock_code + year + reprt_type
        try :
            temp = self.firm_items[a]
            print(temp)
        except :

            R = urllib.request.urlopen(url).read().decode('utf-8')
            D = json.loads(R)
            # print(D)
            if D["status"] == "013" : #== "{'status': '013', 'message': '조회된 데이타가 없습니다.'}":
                # 연결 재무제표를 원하는데, api 상에 자료가 존재하지 않을 경우
                # 단일 재무제표를 대신 다운받음.
                fs_div = "OFS"
                url = "https://opendart.fss.or.kr/api/fnlttSinglAcntAll.json?crtfc_key={}&corp_code={}&bsns_year={}&reprt_code={}&fs_div={}".format(
                    self.key, corp_code, year, code, fs_div)
                R = urllib.request.urlopen(url).read().decode('utf-8')
                D = json.loads(R)  # json 파일 받는 API로 연결받기에...
            if D["status"] == "000":
                print(D)
            else :
                print(D)
            self.firm_items[a] = D


        D = self.firm_items[a]
        thstrm_nm = D["list"][1]['thstrm_nm'][2:D["list"][1]['thstrm_nm'].find(" 기")]
        rcept_no = D["list"][1]['rcept_no'][:8]
        final_dict = dict()
        items = D['list']

        for item in items:
            if item['sj_nm'] == target:
                final_dict[item['account_nm']] = item['thstrm_amount']  # key: 항목 / value: 당기 금액

        if target == "재무상태표":  # 파일 명을 영어로 해야 하기에....
            target = "balance_sheet"
        elif target == "현금흐름표":
            target = "cash_flow"
        elif target == "포괄손익계산서":
            target = "income_statement"

        df = pd.DataFrame(list(final_dict.items()), columns=['항목', '금액']).set_index('항목')
        # print(df)
        title = str(thstrm_nm + "_" + reprt_type + "_" + consolidated + "_" + target + "_" + rcept_no + ".csv")

        sub_folder = "financial_statements"
        temp_path = os.path.join(os.getcwd(), folder,stock_code, sub_folder, target)  # 파일 생성 위한 임시 경로
        self.createDirectory(temp_path)  # 자료 다운로드와 동시에 부모폴더 & 자식폴더 자동 생성
        path = os.path.join(os.getcwd(), folder, stock_code, sub_folder, target, title)  # 데이터 저장할 최종 경로
        df.to_csv(path, encoding="euc-kr")  # 폴더 생성 후 csv 파일로 저장.

    def get_periodic_financial_statements_for_oli(self,firm_name, stock_code, year, reprt_type, consolidated, target):
        if reprt_type == "year":  # 사업보고서
            code = "11011"
        elif reprt_type == "half year":  # 반기보고서
            code = "11012"
        elif reprt_type == "q1":  # 1분기보고서
            code = "11013"
        elif reprt_type == "q3":  # 3분기보고서
            code = "11014"

        corp_code = self.get_corp_code_for_oli(stock_code)  # 기업 코드 가져오기 ---> API 입력용
        # stock_code = self.get_stock_code(firm_name)  # 종목 코드 가져오기 ---> 파일 저장 이름 용

        folder = 'dart'
        if consolidated == "c":
            fs_div = "CFS"  # 연결 재무제표

        elif consolidated == "o":
            fs_div = "OFS"  # 단일 재무제표

        url = "https://opendart.fss.or.kr/api/fnlttSinglAcntAll.json?crtfc_key={}&corp_code={}&bsns_year={}&reprt_code={}&fs_div={}".format(
            self.key, corp_code, year, code, fs_div)

        a = stock_code + year + reprt_type
        try :
            temp = self.firm_items[a]
        except :
            R = urllib.request.urlopen(url).read().decode('utf-8')
            D = json.loads(R)

            if D["status"] == "013" : #== "{'status': '013', 'message': '조회된 데이타가 없습니다.'}":
                # 연결 재무제표를 원하는데, api 상에 자료가 존재하지 않을 경우
                # 단일 재무제표를 대신 다운받음.
                fs_div = "OFS"
                url = "https://opendart.fss.or.kr/api/fnlttSinglAcntAll.json?crtfc_key={}&corp_code={}&bsns_year={}&reprt_code={}&fs_div={}".format(
                    self.key, corp_code, year, code, fs_div)
                R = urllib.request.urlopen(url).read().decode('utf-8')
                D = json.loads(R)  # json 파일 받는 API로 연결받기에...

            if D["status"] == "020":
                print("조회수 초과 :  키변경")
                self.key = "db997885049ef0ea762ad6098e3791d0b6e7c75e"  # henryjeong84268@gmail.com
                fs_div = "CFS"
                url = "https://opendart.fss.or.kr/api/fnlttSinglAcntAll.json?crtfc_key={}&corp_code={}&bsns_year={}&reprt_code={}&fs_div={}".format(
                    self.key, corp_code, year, code, fs_div)
                R = urllib.request.urlopen(url).read().decode('utf-8')
                D = json.loads(R)  # json 파일 받는 API로 연결받기에...

                if D["status"] == "013":  # == "{'status': '013', 'message': '조회된 데이타가 없습니다.'}":
                    # 연결 재무제표를 원하는데, api 상에 자료가 존재하지 않을 경우
                    # 단일 재무제표를 대신 다운받음.
                    fs_div = "OFS"
                    url = "https://opendart.fss.or.kr/api/fnlttSinglAcntAll.json?crtfc_key={}&corp_code={}&bsns_year={}&reprt_code={}&fs_div={}".format(
                        self.key, corp_code, year, code, fs_div)
                    R = urllib.request.urlopen(url).read().decode('utf-8')
                    D = json.loads(R)  # json 파일 받는 API로 연결받기에...

            if D["status"] == "020":
                print("조회수 초과 :  키변경")

                self.key = "604ebad865e4d98c988ff43d37e12c79c887c138"  # pop84268@hotmail.com
                fs_div = "CFS"
                url = "https://opendart.fss.or.kr/api/fnlttSinglAcntAll.json?crtfc_key={}&corp_code={}&bsns_year={}&reprt_code={}&fs_div={}".format(
                    self.key, corp_code, year, code, fs_div)
                R = urllib.request.urlopen(url).read().decode('utf-8')
                D = json.loads(R)  # json 파일 받는 API로 연결받기에...

                if D["status"] == "013":  # == "{'status': '013', 'message': '조회된 데이타가 없습니다.'}":
                    # 연결 재무제표를 원하는데, api 상에 자료가 존재하지 않을 경우
                    # 단일 재무제표를 대신 다운받음.
                    fs_div = "OFS"
                    url = "https://opendart.fss.or.kr/api/fnlttSinglAcntAll.json?crtfc_key={}&corp_code={}&bsns_year={}&reprt_code={}&fs_div={}".format(
                        self.key, corp_code, year, code, fs_div)
                    R = urllib.request.urlopen(url).read().decode('utf-8')
                    D = json.loads(R)  # json 파일 받는 API로 연결받기에...

            if D["status"] == "020":
                print("조회수 초과 :  키변경")

                self.key = "eaf7f1ed3cb845e57fead9cc73b3984698ca36d4"  # hannhyun@gmail.com
                fs_div = "CFS"
                url = "https://opendart.fss.or.kr/api/fnlttSinglAcntAll.json?crtfc_key={}&corp_code={}&bsns_year={}&reprt_code={}&fs_div={}".format(
                    self.key, corp_code, year, code, fs_div)
                R = urllib.request.urlopen(url).read().decode('utf-8')
                D = json.loads(R)  # json 파일 받는 API로 연결받기에...

                if D["status"] == "013":  # == "{'status': '013', 'message': '조회된 데이타가 없습니다.'}":
                    # 연결 재무제표를 원하는데, api 상에 자료가 존재하지 않을 경우
                    # 단일 재무제표를 대신 다운받음.
                    fs_div = "OFS"
                    url = "https://opendart.fss.or.kr/api/fnlttSinglAcntAll.json?crtfc_key={}&corp_code={}&bsns_year={}&reprt_code={}&fs_div={}".format(
                        self.key, corp_code, year, code, fs_div)
                    R = urllib.request.urlopen(url).read().decode('utf-8')
                    D = json.loads(R)  # json 파일 받는 API로 연결받기에...

            if D["status"] == "020":
                print("조회수 초과 :  키변경")

                self.key = "bcafb129d6606e076425808121d7c9a07f5ecabc"  #pop84268@korea.ac.kr
                fs_div = "CFS"
                url = "https://opendart.fss.or.kr/api/fnlttSinglAcntAll.json?crtfc_key={}&corp_code={}&bsns_year={}&reprt_code={}&fs_div={}".format(
                    self.key, corp_code, year, code, fs_div)
                R = urllib.request.urlopen(url).read().decode('utf-8')
                D = json.loads(R)  # json 파일 받는 API로 연결받기에...

                if D["status"] == "013":  # == "{'status': '013', 'message': '조회된 데이타가 없습니다.'}":
                    # 연결 재무제표를 원하는데, api 상에 자료가 존재하지 않을 경우
                    # 단일 재무제표를 대신 다운받음.
                    fs_div = "OFS"
                    url = "https://opendart.fss.or.kr/api/fnlttSinglAcntAll.json?crtfc_key={}&corp_code={}&bsns_year={}&reprt_code={}&fs_div={}".format(
                        self.key, corp_code, year, code, fs_div)
                    R = urllib.request.urlopen(url).read().decode('utf-8')
                    D = json.loads(R)  # json 파일 받는 API로 연결받기에...

            if D["status"] == "020":
                print("조회수 초과 :  키변경")

                self.key = "09e2f1db7f37f6d44f8c6f2b984da99e96c1d4aa"  #jjjang2001@naver.com
                fs_div = "CFS"
                url = "https://opendart.fss.or.kr/api/fnlttSinglAcntAll.json?crtfc_key={}&corp_code={}&bsns_year={}&reprt_code={}&fs_div={}".format(
                    self.key, corp_code, year, code, fs_div)
                R = urllib.request.urlopen(url).read().decode('utf-8')
                D = json.loads(R)  # json 파일 받는 API로 연결받기에...

                if D["status"] == "013":  # == "{'status': '013', 'message': '조회된 데이타가 없습니다.'}":
                    # 연결 재무제표를 원하는데, api 상에 자료가 존재하지 않을 경우
                    # 단일 재무제표를 대신 다운받음.
                    fs_div = "OFS"
                    url = "https://opendart.fss.or.kr/api/fnlttSinglAcntAll.json?crtfc_key={}&corp_code={}&bsns_year={}&reprt_code={}&fs_div={}".format(
                        self.key, corp_code, year, code, fs_div)
                    R = urllib.request.urlopen(url).read().decode('utf-8')
                    D = json.loads(R)  # json 파일 받는 API로 연결받기에...



            if D["status"] == "000":
                # print(D)
                pass

            else :
                pass
                # print(D)
            self.firm_items[a] = D

        D = self.firm_items[a]
        thstrm_nm = D["list"][1]['thstrm_nm'][2:D["list"][1]['thstrm_nm'].find(" 기")]
        rcept_no = D["list"][1]['rcept_no'][:8]
        final_dict = dict()
        items = D['list']
            # print(items)


            # (self,firm_name, stock_code, year, reprt_type, consolidated, target)



        for item in items:
            if item['sj_nm'] == target:
                final_dict[item['account_nm']] = item['thstrm_amount']  # key: 항목 / value: 당기 금액

        if target == "재무상태표":  # 파일 명을 영어로 해야 하기에....
            target = "balance_sheet"
        elif target == "현금흐름표":
            target = "cash_flow"
        elif target == "포괄손익계산서":
            target = "income_statement"

        df = pd.DataFrame(list(final_dict.items()), columns=['항목', '금액']).set_index('항목')
        # print(df)
        title = str(thstrm_nm + "_" + reprt_type + "_" + consolidated + "_" + target + "_" + rcept_no + ".csv")

        sub_folder = "financial_statements"
        temp_path = os.path.join(os.getcwd(), folder, stock_code, sub_folder, target)  # 파일 생성 위한 임시 경로
        self.createDirectory(temp_path)  # 자료 다운로드와 동시에 부모폴더 & 자식폴더 자동 생성
        path = os.path.join(os.getcwd(), folder, stock_code, sub_folder, target, title)  # 데이터 저장할 최종 경로
        df.to_csv(path, encoding="euc-kr")  # 폴더 생성 후 csv 파일로 저장.


if __name__ == "__main__":
    app = QApplication(sys.argv)
    combine_dart = combine_dart()
    combine_dart.show()
    app.exec_()
