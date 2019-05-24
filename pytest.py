# -*- coding: UTF-8 -*-
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd
import xlrd
#import numpy as np
#import pytesseract
import re
#from PIL import Image
import sys
# from PyQt5 import uic, QtWidgets,QtCore
from PyQt5.QtWidgets import QApplication, QDialog, QComboBox, QCheckBox, QListWidget, QLabel, QTextEdit
from PyQt5.QtGui import QPixmap,QIcon
from PyQt5.uic import loadUi
import easygui
import datetime
import sys, os, time
import win32com.client as win32




class Singleton(type):  # type을 상속받음
    _instances = {}  # class의 instance를 저장할 속성

    def __call__(cls, *args, **kwargs):  # class로 instance를 만들 때 호출
        if cls not in cls._instances:  # class로 instance를 생성하지 않았느지 확인
            # 생성하지 않았으면 instance를 생성하여 속성에 저장
            cls._instances[cls] = super(Singleton, cls).__call__(*args, **kwargs)
        return cls._instances[cls]  # class로 instance를 생성했으면 instance 반환


class Dataset(metaclass=Singleton):
    def __init__(self):
        self.row = 0
        self.tabnumber = 0
        self.descriptiondata = ""
        self.breakfastTime = 6
        self.lunchTime = 11
        self.dinnerTime = 16
        self.mealMax = 35000
        self.Entrance = ""
        self.Exit = ""

        self.Pandas_header = 4
        self.Pandas_CostIndex = '거래금액'
        self.Pandas_Cost_isStr = False
        self.Pandas_DateIndex = '거래일자'
        self.Pandas_TimeIndex = '거래시간'
        self.Pandas_DateTimeIndex = '거래일시'
        self.Pandas_ApprovalNumIndex = '승인번호'
        self.Pandas_ApprovalNum_isStr = False
        self.Pandas_TaxiBusIndex1 = '가맹점명'
        self.Pandas_TaxiBusIndex2 = '가맹점업종'
        self.Pandas_DateFormat = '%Y/%m/%d'
        self.Pandas_TimeFormat = '%H:%M:%S'
        self.Pandas_DateTimeCombined = True
        self.Pandas_DateTimeFormat = '%Y/%m/%d %H:%M:%S'
        self.Pandas_TimelocData = 0, 19
        self.Pandas_DatelocData = 0, 20
        self.Pandas_Entrance = '입구'
        self.Pandas_Exit = '출구'

        self.FileDirectory = ""

    def seteightnumbers(self, data):
        self.eightnumbers = data

    def geteightnumbers(self):
        return self.eightnumbers

    def setEnvelope(self, data):
        self.Envelope = data

    def getEnvelope(self):
        return self.Envelope

    def setDate(self, data):
        self.Date = data

    def getDate(self):
        return self.Date

    def setcost(self, data):
        self.cost = data

    def getcost(self):
        return self.cost

    def setreceiptimageAddress(self, data):
        self.receiptimageAddress = data

    def getreceiptimageAddress(self):
        return self.receiptimageAddress

    def setrow(self, data):
        self.row = data

    def getrow(self):
        return self.row

    def setexcelAddress(self, data):
        self.excelAddress = data

    def getexcelAddress(self):
        return self.excelAddress

    def setexcelData(self, data):
        self.excelData = data

    def getexcelData(self):
        return self.excelData

    def setbrowser(self, data):
        self.browser = data

    def getbrowser(self):
        return self.browser

    def setInvoicenum(self, data):
        self.Invoicenum = data

    def getInvoicenum(self):
        return self.Invoicenum

    def settabnumber(self, data):
        self.tabnumber = data

    def gettabnumber(self):
        return self.tabnumber

    def setbreakfastTime(self, data):
        self.breakfastTime = data

    def getbreakfastTime(self):
        return self.breakfastTime

    def setlunchTime(self, data):
        self.lunchTime = data

    def getlunchTime(self):
        return self.lunchTime

    def setdinnerTime(self, data):
        self.dinnerTime = data

    def getdinnerTime(self):
        return self.dinnerTime

    def setmealMax(self, data):
        self.mealMax = data

    def getmealMax(self):
        return self.mealMax

    def setmeal(self, data):
        self.meal = data

    def getmeal(self):
        return self.meal

    def setdescriptiondata(self, data):
        self.descriptiondata = data

    def getdescriptiondata(self):
        return self.descriptiondata

    def setCardType(self, data):
        self.CardType = data

    def getCardType(self):
        return self.CardType

    def setFileDirectory(self, data):
        self.FileDirectory = data

    def getFileDirectory(self):
        return self.FileDirectory

    def setLastDate(self, data):
        self.LastDate = data

    def getLastDate(self):
        return self.LastDate

    def setUiMain(self, data):
        self.UiMain = data

    def getUiMain(self):
        return self.UiMain

    def setUiReference(self, data):
        self.UiReference = data

    def getUiReference(self):
        return self.UiReference

    def setABBIcon(self, data):
        self.ABBIcon = data

    def getABBIcon(self):
        return self.ABBIcon

    def setOnebyOne(self, data):
        self.OnebyOne = data

    def getOnebyOne(self):
        return self.OnebyOne

    def setEntrance(self, data):
        self.Entrance = data

    def getEntrance(self):
        return self.Entrance

    def setExit(self, data):
        self.Exit = data

    def getExit(self):
        return self.Exit

    def CardDataAssignment(self):
        if self.getCardType() == 'ABB Card - Meal & Accommodation':
            self.Pandas_header = 1
            self.Pandas_CostIndex = '거래금액'
            self.Pandas_Cost_isStr = True
            self.Pandas_DateIndex = '거래일자'
            self.Pandas_TimeIndex = '거래시간'
            #self.Pandas_DateTimeIndex = '거래일시'
            self.Pandas_ApprovalNumIndex = '승인번호'
            self.Pandas_ApprovalNum_isStr = False
            self.Pandas_TaxiBusIndex1 = '가맹점명'
            self.Pandas_TaxiBusIndex2 = '가맹점업종'
            self.Pandas_DateFormat = '%Y-%m-%d'
            self.Pandas_TimeFormat = '%H:%M:%S'
            self.Pandas_DateTimeCombined = False
            #self.Pandas_DateTimeFormat = '%Y/%m/%d %H:%M:%S'
            self.Pandas_DatelocData = 0, 5
            self.Pandas_TimelocData = 0, 6
            #self.Pandas_Entrance = '입구'
            #self.Pandas_Exit = '출구'

        elif self.getCardType() == 'Hana Card - Meal & Accommodation':
            self.Pandas_header = 3
            self.Pandas_CostIndex = '승인금액'
            self.Pandas_Cost_isStr = True
            self.Pandas_DateIndex = '이용일'
            self.Pandas_TimeIndex = '이용시간'
            #self.Pandas_DateTimeIndex = '거래일시'
            self.Pandas_ApprovalNumIndex = '승인번호'
            self.Pandas_ApprovalNum_isStr = False
            # self.Pandas_TaxiBusIndex1 = '가맹점명'
            # self.Pandas_TaxiBusIndex2 = '가맹점업종'
            self.Pandas_DateFormat = '%Y.%m.%d'
            self.Pandas_TimeFormat = '%H:%M:%S'
            self.Pandas_DateTimeCombined = False
            #self.Pandas_DateTimeFormat = '%Y/%m/%d %H:%M:%S'
            self.Pandas_DatelocData = 0, 0
            self.Pandas_TimelocData = 0, 3
            #self.Pandas_Entrance = '입구'
            #self.Pandas_Exit = '출구'

        elif self.getCardType() == 'Toll Gate':
            self.Pandas_header = 4
            self.Pandas_CostIndex = '거래금액'
            self.Pandas_Cost_isStr = False
            self.Pandas_DateIndex = '거래일자'
            self.Pandas_TimeIndex = '거래시간'
            self.Pandas_DateTimeIndex = '거래일시'
            self.Pandas_ApprovalNumIndex = '승인번호'
            self.Pandas_ApprovalNum_isStr = False
            # self.Pandas_TaxiBusIndex1 = '가맹점명'
            # self.Pandas_TaxiBusIndex2 = '가맹점업종'
            self.Pandas_DateFormat = '%Y-%m-%d'
            self.Pandas_TimeFormat = '%H:%M:%S'
            self.Pandas_DateTimeCombined = True #거래일자랑 거래시간이 하나임
            self.Pandas_DateTimeFormat = '%Y/%m/%d %H:%M:%S'
            self.Pandas_DatelocData = 0, 19
            self.Pandas_TimelocData = 0, 20
            self.Pandas_Entrance = '입구'
            self.Pandas_Exit = '출구'
        else:
            print('Error!')



class MealCalculation(metaclass=Singleton):
    def __init__(self):
        self.dataset = Dataset()
        self.chromefilling = ChromeFilling()

    def start(self):
        #exceldata = pd.read_excel("C:\\Users\\krjikim\\Downloads\\카드승인내역_1548996372136_20190201.xls", header=1)
        exceldata = pd.read_excel(self.dataset.getexcelAddress(), header=self.dataset.Pandas_header)
        #exceldata = pd.read_excel(self.dataset.getexcelData(), header=self.dataset.Pandas_header)
        print(exceldata)
        if self.dataset.Pandas_Cost_isStr == True:
            exceldata[self.dataset.Pandas_CostIndex] = exceldata[self.dataset.Pandas_CostIndex].str.replace(',', '')
        # print(exceldata.loc[(exceldata[self.dataset.Pandas_CostIndex].str.contains('-')),:])
            exceldata[self.dataset.Pandas_CostIndex] = pd.to_numeric(exceldata[self.dataset.Pandas_CostIndex])

        if self.dataset.Pandas_ApprovalNum_isStr == True:
            exceldata[self.dataset.Pandas_ApprovalNumIndex] = pd.to_numeric(exceldata[self.dataset.Pandas_ApprovalNumIndex])
        exceldata = exceldata.drop(exceldata[(exceldata[self.dataset.Pandas_CostIndex]) <= 0].index)
        if self.dataset.getCardType() == 'ABB Card - Meal & Accommodation':
            exceldata.drop(exceldata.loc[(exceldata[self.dataset.Pandas_TaxiBusIndex1].str.contains('택시'))].index, inplace=True)
            exceldata.drop(exceldata.loc[(exceldata[self.dataset.Pandas_TaxiBusIndex2].str.contains('택시|버스|기차'))].index, inplace=True)
        exceldata.reset_index(drop=True, inplace=True)
        print(exceldata)
        if self.dataset.Pandas_DateTimeCombined == False:
            exceldata[self.dataset.Pandas_DateIndex] = pd.to_datetime(exceldata[self.dataset.Pandas_DateIndex], format=self.dataset.Pandas_DateFormat, errors='coerce').dt.date
            exceldata[self.dataset.Pandas_TimeIndex] = pd.to_datetime(exceldata[self.dataset.Pandas_TimeIndex], format=self.dataset.Pandas_TimeFormat, errors='coerce').dt.time
        elif self.dataset.Pandas_DateTimeCombined == True:
            exceldata[self.dataset.Pandas_DateTimeIndex] = pd.to_datetime(exceldata[self.dataset.Pandas_DateTimeIndex], format=self.dataset.Pandas_DateTimeFormat)
            exceldata[self.dataset.Pandas_DateIndex] = pd.to_datetime(exceldata[self.dataset.Pandas_DateTimeIndex], format=self.dataset.Pandas_DateFormat, errors='coerce').dt.date
            exceldata[self.dataset.Pandas_TimeIndex] = pd.to_datetime(exceldata[self.dataset.Pandas_DateTimeIndex], format=self.dataset.Pandas_TimeFormat, errors='coerce').dt.time

        # print(exceldata)
        self.dategroup = exceldata.groupby(self.dataset.Pandas_DateIndex, sort=False)
        timedivision1 = datetime.time(int(self.dataset.breakfastTime))
        timedivision2 = datetime.time(int(self.dataset.lunchTime))
        timedivision3 = datetime.time(int(self.dataset.dinnerTime))
        costMax = int(self.dataset.getmealMax())
        nextdinner=0
        nexttoll=0
        nexttoll_nodate=0
        nextaccommodation=0
        nextdinner_nodate=0
        nextaccommodation_nodate=0

        for dategroupdata in self.dategroup:
            descriptiondata=[]
            breakfast = 0
            lunch = 0
            dinner = 0
            cost = 0
            accommodation = 0
            invoicenum=0
            count=-1

            if nexttoll!=0:
                cost+=nexttoll
                invoicenum = None
                dinner_rightaway=True
                nexttoll=0

            if nextdinner!=0:
                dinner+=nextdinner
                cost+=nextdinner
                invoicenum = invoicenum_nextday
                dinner_rightaway=True
                nextdinner=0
                invoicenum_nextday=0

            if nextaccommodation != 0:
                accommodation = nextaccommodation
                accommodation = int(dategroupdata[1][self.dataset.Pandas_CostIndex].loc[indexdata + 1])
                invoicenum_accommodation = int(dategroupdata[1][self.dataset.Pandas_ApprovalNumIndex].loc[indexdata + 1])
                total = 'Total : ' + str(accommodation)
                descriptiondata.append(total)
                print('Accommodation : ' + str(accommodation))
                self.dataset.setEnvelope('Accommodation')
                self.dataset.setcost(accommodation)
                self.dataset.setInvoicenum(invoicenum_nextday_accommodation)
                self.dataset.setdescriptiondata(descriptiondata)
                nextaccommodation = 0
                invoicenum_nextday_accommodation = 0
                self.form_reference.start()
                self.chromefilling.start()
                self.form_reference.workdone()


            for timedata in dategroupdata[1][self.dataset.Pandas_TimeIndex]:
                indexdata = dategroupdata[1][self.dataset.Pandas_TimeIndex][dategroupdata[1][self.dataset.Pandas_TimeIndex] == timedata].index.values
                count+=1

                if self.dataset.getLastDate() != (datetime.date(2000,1,1)) and exceldata.loc[indexdata].iloc[self.dataset.Pandas_DatelocData] != self.dataset.getLastDate():
                    print(self.dataset.getLastDate())
                    print(type(self.dataset.getLastDate()))
                #invoicenum = int(dategroupdata[1][self.dataset.Pandas_ApprovalNumIndex].loc[indexdata])
                else:
                    self.dataset.setLastDate(datetime.date(2000,1,1))
                    if self.dataset.getCardType() == 'Toll Gate':
                        if exceldata.loc[indexdata].iloc[self.dataset.Pandas_TimelocData] < timedivision1 and exceldata.loc[indexdata + len(dategroupdata[1])-count].iloc[self.dataset.Pandas_DatelocData] == (exceldata.loc[indexdata].iloc[self.dataset.Pandas_DatelocData] - datetime.timedelta(days=1)):  # 새벽인데 전날 출장 있을경우
                            nexttoll += int(dategroupdata[1][self.dataset.Pandas_CostIndex].loc[indexdata])
                        elif exceldata.loc[indexdata].iloc[self.dataset.Pandas_TimelocData] < timedivision1 and exceldata.loc[indexdata + len(dategroupdata[1])-count].iloc[self.dataset.Pandas_DatelocData] != (exceldata.loc[indexdata].iloc[self.dataset.Pandas_DatelocData] - datetime.timedelta(days=1)):  # 새벽인데 전날 출장 없을경우
                            nexttoll_nodate += int(dategroupdata[1][self.dataset.Pandas_CostIndex].loc[indexdata])
                        else:
                            cost += int(dategroupdata[1][self.dataset.Pandas_CostIndex].loc[indexdata])
                        #print(exceldata[self.dataset.Pandas_Entrance].loc[indexdata].reset_index(drop=True))
                        self.dataset.setEntrance(exceldata[self.dataset.Pandas_Entrance].loc[indexdata].reset_index(drop=True)[0])
                        self.dataset.setExit(exceldata[self.dataset.Pandas_Exit].loc[indexdata].reset_index(drop=True)[0])
                    else:
                        if exceldata.loc[indexdata].iloc[self.dataset.Pandas_TimelocData] < timedivision1 and exceldata.loc[indexdata + len(dategroupdata[1])-count].iloc[self.dataset.Pandas_DatelocData] == (exceldata.loc[indexdata].iloc[self.dataset.Pandas_DatelocData] - datetime.timedelta(days=1)):  # 새벽인데 전날 출장 있을경우
                            if int(dategroupdata[1][self.dataset.Pandas_CostIndex].loc[indexdata]) < costMax:
                                nextdinner += int(dategroupdata[1][self.dataset.Pandas_CostIndex].loc[indexdata])
                                invoicenum_nextday = int(dategroupdata[1][self.dataset.Pandas_ApprovalNumIndex].loc[indexdata])
                            else:
                                nextaccommodation += int(dategroupdata[1][self.dataset.Pandas_CostIndex].loc[indexdata])
                                invoicenum_nextday_accommodation = int(dategroupdata[1][self.dataset.Pandas_ApprovalNumIndex].loc[indexdata])
                        elif exceldata.loc[indexdata].iloc[self.dataset.Pandas_TimelocData] < timedivision1 and exceldata.loc[indexdata + len(dategroupdata[1])-count].iloc[self.dataset.Pandas_DatelocData] != (exceldata.loc[indexdata].iloc[self.dataset.Pandas_DatelocData] - datetime.timedelta(days=1)):  # 새벽인데 전날 출장 없을경우
                            if int(dategroupdata[1][self.dataset.Pandas_CostIndex].loc[indexdata]) < costMax:
                                nextdinner_nodate += int(dategroupdata[1][self.dataset.Pandas_CostIndex].loc[indexdata])
                                invoicenum_nextday = int(dategroupdata[1][self.dataset.Pandas_ApprovalNumIndex].loc[indexdata])
                            else:
                                nextaccommodation_nodate = int(dategroupdata[1][self.dataset.Pandas_CostIndex].loc[indexdata])
                                invoicenum_nextday_accommodation = int(dategroupdata[1][self.dataset.Pandas_ApprovalNumIndex].loc[indexdata])
                        elif timedivision1 < timedata < timedivision2:
                            if int(dategroupdata[1][self.dataset.Pandas_CostIndex].loc[indexdata]) < costMax:
                                breakfast += int(dategroupdata[1][self.dataset.Pandas_CostIndex].loc[indexdata])
                                cost += int(dategroupdata[1][self.dataset.Pandas_CostIndex].loc[indexdata])
                                invoicenum = int(dategroupdata[1][self.dataset.Pandas_ApprovalNumIndex].loc[indexdata])
                            else:
                                accommodation = int(dategroupdata[1][self.dataset.Pandas_CostIndex].loc[indexdata])
                                invoicenum_accommodation = int(dategroupdata[1][self.dataset.Pandas_ApprovalNumIndex].loc[indexdata])
                        elif timedivision2 < timedata < timedivision3:
                            if int(dategroupdata[1][self.dataset.Pandas_CostIndex].loc[indexdata]) < costMax:
                                lunch += int(dategroupdata[1][self.dataset.Pandas_CostIndex].loc[indexdata])
                                cost += int(dategroupdata[1][self.dataset.Pandas_CostIndex].loc[indexdata])
                                invoicenum = int(dategroupdata[1][self.dataset.Pandas_ApprovalNumIndex].loc[indexdata])
                            else:
                                accommodation = int(dategroupdata[1][self.dataset.Pandas_CostIndex].loc[indexdata])
                                invoicenum_accommodation = int(dategroupdata[1][self.dataset.Pandas_ApprovalNumIndex].loc[indexdata])
                        elif timedivision3 < timedata:
                            if int(dategroupdata[1][self.dataset.Pandas_CostIndex].loc[indexdata]) < costMax:
                                dinner += int(dategroupdata[1][self.dataset.Pandas_CostIndex].loc[indexdata])
                                cost += int(dategroupdata[1][self.dataset.Pandas_CostIndex].loc[indexdata])
                                invoicenum = int(dategroupdata[1][self.dataset.Pandas_ApprovalNumIndex].loc[indexdata])
                            else:
                                accommodation = int(dategroupdata[1][self.dataset.Pandas_CostIndex].loc[indexdata])
                                invoicenum_accommodation = int(dategroupdata[1][self.dataset.Pandas_ApprovalNumIndex].loc[indexdata])
                        else:
                            print('Error!!')
                if self.dataset.getOnebyOne():
                    if breakfast != 0:
                        breakfastdata = 'Breakfast : ' + str(breakfast)
                        print(breakfastdata)
                        descriptiondata.append(breakfastdata)
                    if lunch != 0:
                        lunchdata = 'Lunch : ' + str(lunch)
                        print(lunchdata)
                        descriptiondata.append(lunchdata)
                    if dinner != 0:
                        dinnerdata = 'Dinner : ' + str(dinner)
                        print(dinnerdata)
                        descriptiondata.append(dinnerdata)
                    if accommodation != 0:
                        #accommodationdata = str(accommodation)
                        #descriptiondata.append(accommodationdata)
                        cost = accommodation
                    if cost != 0:
                        total = 'Total : ' + str(cost)
                        if self.dataset.getCardType() == 'Toll Gate':
                            self.dataset.setEnvelope('Toll Gate')
                            #descriptiondata.append(self.dataset.getExit())
                        elif accommodation != 0:
                            self.dataset.setEnvelope('Accommodation')
                            self.dataset.setInvoicenum(invoicenum_accommodation)
                        else:
                            self.dataset.setEnvelope('Meal')
                            self.dataset.setInvoicenum(invoicenum)
                        print(total)
                        descriptiondata.append(total)
                        self.dataset.setDate(str(dategroupdata[0] + datetime.timedelta(days=1)))
                        self.dataset.setcost(cost)
                        self.dataset.setdescriptiondata(descriptiondata)
                        self.form_reference = Form_reference()
                        self.form_reference.start()
                        self.chromefilling.start()
                        self.form_reference.workdone()
                        descriptiondata = []
                        breakfast = 0
                        lunch = 0
                        dinner = 0
                        cost = 0
                        accommodation = 0
                        invoicenum = 0

            if self.dataset.getOnebyOne()==False:
                if breakfast != 0:
                    breakfastdata='Breakfast : ' + str(breakfast)
                    print(breakfastdata)
                    descriptiondata.append(breakfastdata)
                if lunch != 0:
                    lunchdata='Lunch : ' + str(lunch)
                    print(lunchdata)
                    descriptiondata.append(lunchdata)
                if dinner != 0:
                    dinnerdata='Dinner : ' + str(dinner)
                    print(dinnerdata)
                    descriptiondata.append(dinnerdata)
                if cost!=0:
                    total='Total : ' + str(cost)
                    if self.dataset.getCardType() == 'Toll Gate':
                        self.dataset.setEnvelope('Toll Gate')
                        descriptiondata.append(self.dataset.getEntrance()+'<->'+self.dataset.getExit())
                    else:
                        self.dataset.setEnvelope('Meal')
                        self.dataset.setInvoicenum(invoicenum)
                    print(total)
                    descriptiondata.append(total)
                    self.dataset.setDate(str(dategroupdata[0]+datetime.timedelta(days=1)))
                    self.dataset.setcost(cost)
                    self.dataset.setdescriptiondata(descriptiondata)
                    self.form_reference = Form_reference()
                    self.form_reference.start()
                    self.chromefilling.start()
                    self.form_reference.workdone()


                    if accommodation != 0:
                        descriptiondata = []
                        total = 'Total : ' + str(accommodation)
                        descriptiondata.append(total)
                        print('Accommodation : ' + str(accommodation))
                        self.dataset.setEnvelope('Accommodation')
                        self.dataset.setcost(accommodation)
                        self.dataset.setInvoicenum(invoicenum_accommodation)
                        self.dataset.setdescriptiondata(descriptiondata)
                        self.form_reference.start()
                        self.chromefilling.start()
                        self.form_reference.workdone()


                    if nextdinner_nodate!=0 or nextaccommodation_nodate!=0 or nexttoll_nodate:
                        self.dataset.setDate(str(dategroupdata[0]))
                        if nextdinner_nodate!=0:
                            descriptiondata = []
                            dinnerdata = 'Dinner : ' + str(nextdinner_nodate)
                            print(dinnerdata)
                            descriptiondata.append(dinnerdata)
                            total = 'Total : ' + str(nextdinner_nodate)
                            descriptiondata.append(total)
                            print('Total_nextday : ' + str(nextdinner_nodate))
                            self.dataset.setEnvelope('Meal')
                            self.dataset.setcost(nextdinner_nodate)
                            self.dataset.setInvoicenum(invoicenum_nextday)
                            self.dataset.setdescriptiondata(descriptiondata)
                            self.form_reference.start()
                            self.chromefilling.start()
                            self.form_reference.workdone()
                            nextdinner_nodate =0
                        if nextaccommodation_nodate!=0:
                            descriptiondata = []
                            total = 'Total : ' + str(nextaccommodation_nodate)
                            descriptiondata.append(total)
                            print('Accommodation : ' + str(nextaccommodation_nodate))
                            self.dataset.setEnvelope('Accommodation')
                            self.dataset.setcost(nextaccommodation_nodate)
                            self.dataset.setInvoicenum(invoicenum_nextday_accommodation)
                            self.dataset.setdescriptiondata(descriptiondata)
                            self.form_reference.start()
                            self.chromefilling.start()
                            self.form_reference.workdone()
                            nextaccommodation_nodate =0
                        if nexttoll_nodate!=0:
                            descriptiondata = []
                            tolldata = 'Toll : ' + str(nexttoll_nodate)
                            print(tolldata)
                            descriptiondata.append(tolldata)
                            total = 'Total : ' + str(nexttoll_nodate)
                            descriptiondata.append(total)
                            print('Total_nextday : ' + str(nexttoll_nodate))
                            self.dataset.setEnvelope('Toll')
                            self.dataset.setcost(nexttoll_nodate)
                            self.dataset.setdescriptiondata(descriptiondata)
                            self.form_reference.start()
                            self.chromefilling.start()
                            self.form_reference.workdone()
                            nextdinner_nodate = 0

#
#
# class tesseractjob(metaclass=Singleton):
#     def __init__(self):
#         self.dataset = Dataset()
#     def start(self):
#         pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"
#         try:
#             img = Image.open(self.dataset.getreceiptimageAddress())
#             text = pytesseract.image_to_string(img, lang='eng3')
#             p = re.compile('\d\d\d\d\d\d\d\d')
#             self.dataset.seteightnumbers(p.findall(text))
#             print(self.dataset.geteightnumbers())
#         except Exception as e:
#             print(e)


class ChromeFilling(metaclass=Singleton):
    def __init__(self):
        self.dataset = Dataset()

    def start(self):

        self.dataset.setreceiptimageAddress(easygui.fileopenbox(default=self.dataset.getFileDirectory()))

        regex = r".+\\"
        print(type(re.findall(regex,self.dataset.getreceiptimageAddress())[0]))
        self.dataset.setFileDirectory(re.findall(regex,self.dataset.getreceiptimageAddress())[0])

        self.dataset.getbrowser().get('https://ete.abb.com.cn/Invoice/Detail')

        Envelope = self.dataset.getbrowser().find_element_by_class_name("k-input")
        Envelope.send_keys(self.dataset.getEnvelope())

        InvoiceDate = self.dataset.getbrowser().find_element_by_id("upload")
        InvoiceDate.send_keys(self.dataset.getreceiptimageAddress())

        Amount = self.dataset.getbrowser().find_element_by_class_name("k-formatted-value")
        #Amount = self.dataset.getbrowser().find_elements_by_xpath("//input[@class='k-formatted-value k-input']")[0]
        Amount.send_keys(self.dataset.getcost())

        InvoiceDate = self.dataset.getbrowser().find_element_by_id("InvoiceDate")
        InvoiceDate.clear()
        InvoiceDate.send_keys(self.dataset.getDate())

        if self.dataset.getCardType() != 'Toll Gate':
            InvoiceNo = self.dataset.getbrowser().find_element_by_id("InvoiceNO")
            InvoiceNo.send_keys(self.dataset.getInvoicenum())

        InvoiceCode = self.dataset.getbrowser().find_element_by_id("InvoiceCode")
        InvoiceCode.send_keys(self.dataset.getEnvelope())

        InvoiceCode = self.dataset.getbrowser().find_element_by_id("Description")
        for items in self.dataset.getdescriptiondata():
            if type(items) == str:
                InvoiceCode.send_keys(str(items) + "\n")

        self.dataset.settabnumber(self.dataset.gettabnumber()+1)
        self.dataset.getbrowser().execute_script("window.open();")
        self.dataset.getbrowser().switch_to_window(self.dataset.getbrowser().window_handles[self.dataset.gettabnumber()])



class Form_reference(QDialog):
    def __init__(self):
        super().__init__()
        self.dataset = Dataset()
        try:
            self.dataset.setUiReference(os.path.join(sys._MEIPASS,"Reference.ui"))
        except Exception:
            self.dataset.setUiReference(os.path.join(os.path.abspath("."),"Reference.ui"))
        loadUi(self.dataset.getUiReference(), self)
        self.setWindowIcon(QIcon(self.dataset.getABBIcon()))
        self.move(1200,450)
        self.show()
    def start(self):
        self.listWidget.addItem(str(datetime.datetime.strptime(self.dataset.getDate(),'%Y-%m-%d').date()-datetime.timedelta(days=1)))
        descriptiondata = self.dataset.getdescriptiondata()
        for i in descriptiondata:
            self.listWidget.addItem(i)
    def workdone(self):
        item = self.listWidget.clear()

class Form_base(QDialog, QComboBox,QLabel,QIcon):
    def __init__(self):
        super().__init__()
        #self.tesseract = tesseractjob()
        self.dataset = Dataset()
        self.chromefilling = ChromeFilling()
        self.mealcalculation = MealCalculation()
        try:
            self.dataset.setABBIcon(os.path.join(sys._MEIPASS,"ABBMark.ico"))
            ABBMark = os.path.join(sys._MEIPASS,"ABBMark.jpg")
            MealTimeline= os.path.join(sys._MEIPASS,"MealTimeline.png")
            self.dataset.setUiMain(os.path.join(sys._MEIPASS,"Main.ui"))
        except Exception:
            self.dataset.setABBIcon(os.path.join(os.path.abspath("."),"ABBMark.ico"))
            ABBMark= os.path.join(os.path.abspath("."),"ABBMark.jpg")
            MealTimeline= os.path.join(os.path.abspath("."),"MealTimeline.png")
            self.dataset.setUiMain(os.path.join(os.path.abspath("."),"Main.ui"))
        loadUi(self.dataset.getUiMain(), self)
        self.setWindowIcon(QIcon(self.dataset.getABBIcon()))
        self.pushButton.clicked.connect(self.pushButton_clicked1)
        #self.pushButton_2.clicked.connect(self.pushButton_clicked2)
        self.pushButton_3.clicked.connect(self.pushButton_clicked3)
        self.pushButton_4.clicked.connect(self.pushButton_clicked4)
        self.pushButton_5.clicked.connect(self.pushButton_clicked5)
        self.label.setPixmap(QPixmap(ABBMark))
        self.label_2.setPixmap(QPixmap(MealTimeline))

    def pushButton_clicked1(self):
        try:
            #self.chromefilling.start()
            #self.tesseract.start()
            #self.filling.exec()
            self.dataset.setbreakfastTime(self.textEdit_1.toPlainText())
            self.dataset.setlunchTime(self.textEdit_2.toPlainText())
            self.dataset.setdinnerTime(self.textEdit_3.toPlainText())
            self.dataset.setmealMax(self.textEdit_4.toPlainText())
            #QCheckBox.isChecked()

            if self.checkBox.isChecked():
                #self.dataset.setSpecificBox_isChecked(True)
                year = self.textEdit_5.toPlainText()
                month = self.textEdit_6.toPlainText()
                day = self.textEdit_7.toPlainText()
                self.dataset.setLastDate(datetime.date(int(year),int(month),int(day)))
            else:
                #self.dataset.setSpecificBox_isChecked(False)
                self.dataset.setLastDate(datetime.date(2000,1,1))

            if  self.checkBox_2.isChecked():
                self.dataset.setOnebyOne(True)
            else:
                self.dataset.setOnebyOne(False)
            self.dataset.setCardType(self.comboBox.currentText())
            self.dataset.CardDataAssignment()
            self.mealcalculation.start()
        except Exception as e:
            print(e)

    # def pushButton_clicked2(self):
    #     self.dataset.setreceiptimageAddress(easygui.fileopenbox())
    #     print(self.dataset.getreceiptimageAddress())

    def pushButton_clicked3(self):
        try:
            fname = easygui.fileopenbox()
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.DisplayAlerts = False
            wb = excel.Workbooks.Open(fname)

            wb.SaveAs(fname + "x", FileFormat=51)  # FileFormat = 51 is for .xlsx extension
            wb.Close()  # FileFormat = 56 is for .xls extension
            excel.Application.Quit()
            #excel = pd.read_excel(self.dataset.getexcelAddress(), header=self.dataset.Pandas_header)
            #print(self.dataset.getexcelData())
            print(fname+'x')
            self.dataset.setexcelAddress(fname+'x')
        except Exception as e:
            print(e)
        #self.dataset.setexcelData(pd.read_excel(self.dataset.getexcelAddress()))
        #print(self.dataset.getexcelAddress())


    def pushButton_clicked4(self):
        # data = pd.read_excel("C:\\Users\\krjikim\\Desktop\\Jin\\Expense\\Expense2.xlsx")

        opts = webdriver.ChromeOptions()
        opts.add_experimental_option("detach", True)
        if getattr(sys, 'frozen', False):
            chromedriver_path = os.path.join(sys._MEIPASS, "chromedriver.exe")
            chromedriver = webdriver.Chrome(chromedriver_path,chrome_options=opts)
        else:
            chromedriver = webdriver.Chrome(chrome_options=opts)

        #chromedriver = 'chromedriver.exe'
        self.dataset.setbrowser(chromedriver)
        self.dataset.getbrowser().get('https://ete.abb.com.cn/Login/Login_classic')


    def pushButton_clicked5(self):
        self.dataset.settabnumber(0)

    def listWidgetChange(self):
        self.form_base.listWidget = QListWidget()
        for i in self.dataset.getdescriptiondata():
            self.form_base.listWidget.addItem(i)


#
# class Form(QDialog):
#     def __init__(self):
#         super().__init__()
#         self.dataset = Dataset()
#         self.tesseract = tesseractjob()
#         self.chromeFilling = ChromeFilling()
#         loadUi("C:\\Users\\krjikim\\Desktop\\Jin\\Study\\test1.ui", self)
#         self.pushButton.clicked.connect(self.pushButton_clicked1)
#         self.pushButton_2.clicked.connect(self.pushButton_clicked2)
#         self.pushButton_3.clicked.connect(self.pushButton_clicked3)
#         self.pushButton_4.clicked.connect(self.pushButton_clicked4)
#         self.pushButton_5.clicked.connect(self.pushButton_clicked5)
#         self.pushButton_6.clicked.connect(self.pushButton_clicked6)
#         self.pushButton_7.clicked.connect(self.pushButton_clicked7)
#         self.pushButton_8.clicked.connect(self.pushButton_clicked8)
#         self.pushButton_9.clicked.connect(self.pushButton_clicked9)
#         self.pushButton_10.clicked.connect(self.pushButton_clicked10)
#         self.pushButton_11.clicked.connect(self.pushButton_clicked11)
#         self.m = self.dataset.geteightnumbers()
#         try:
#             self.pushButton.setText(self.m[0])
#             self.pushButton_2.setText(self.m[1])
#             self.pushButton_3.setText(self.m[2])
#             self.pushButton_4.setText(self.m[3])
#             self.pushButton_5.setText(self.m[4])
#             self.pushButton_6.setText(self.m[5])
#             self.pushButton_7.setText(self.m[6])
#             self.pushButton_8.setText(self.m[7])
#             self.pushButton_9.setText(self.m[8])
#             self.pushButton_10.setText(self.m[9])
#         except Exception as e:
#             print(e)
#             pass
#         self.show()
#
#     def pushButton_clicked1(self):
#         try:
#             invoice = self.m[0]
#         except Exception as e:
#             print(e)
#         self.textEdit.setText(invoice)
#
#     def pushButton_clicked2(self):
#         try:
#             invoice = self.m[1]
#         except Exception as e:
#             print(e)
#         self.textEdit.setText(invoice)
#
#     def pushButton_clicked3(self):
#         try:
#             invoice = self.m[2]
#         except Exception as e:
#             print(e)
#         self.textEdit.setText(invoice)
#
#     def pushButton_clicked4(self):
#         try:
#             invoice = self.m[3]
#         except Exception as e:
#             print(e)
#         self.textEdit.setText(invoice)
#
#     def pushButton_clicked5(self):
#         try:
#             invoice = self.m[4]
#         except Exception as e:
#             print(e)
#         self.textEdit.setText(invoice)
#
#     def pushButton_clicked6(self):
#         try:
#             invoice = self.m[5]
#         except Exception as e:
#             print(e)
#         self.textEdit.setText(invoice)
#
#     def pushButton_clicked7(self):
#         try:
#             invoice = self.m[6]
#         except Exception as e:
#             print(e)
#         self.textEdit.setText(invoice)
#
#     def pushButton_clicked8(self):
#         try:
#             invoice = self.m[7]
#         except Exception as e:
#             print(e)
#         self.textEdit.setText(invoice)
#
#     def pushButton_clicked9(self):
#         try:
#             invoice = self.m[8]
#         except Exception as e:
#             print(e)
#         self.textEdit.setText(invoice)
#
#     def pushButton_clicked10(self):
#         try:
#             invoice = self.m[9]
#         except Exception as e:
#             print(e)
#         self.textEdit.setText(invoice)
#
#     def pushButton_clicked11(self):
#         try:
#             self.dataset.setinvoicenum(self.textEdit.toPlainText())
#             print(self.dataset.getinvoicenum())
#             self.close()
#             print(self.dataset.getrow())
#             self.chromeFilling.start()
#         except Exception as e:
#             print(e)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = Form_base()
    w.show()
    sys.exit(app.exec())