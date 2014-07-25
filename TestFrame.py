# -*- coding:utf-8 -*-

#****************************************************************
# TestFrame.py
# Author     : kiven
# Version    : 1.1.2
# Date       : 2014-07-21
# Description: 自动化测试平台
#****************************************************************

import os,sys, urllib, httplib, profile, datetime, time
from xml2dict import XML2Dict
import win32com.client
from win32com.client import Dispatch
import xml.etree.ElementTree as et
#import MySQLdb

#Excel表格中测试结果底色
OK_COLOR=0xffffff
NG_COLOR=0xff
#NT_COLOR=0xffff
NT_COLOR=0xC0C0C0

#Excel表格中测试结果汇总显示位置
TESTTIME=[1, 14]
TESTRESULT=[2, 14]

#Excel模版设置
#self.titleindex=3        #Excel中测试用例标题行索引
#self.casebegin =4        #Excel中测试用例开始行索引
#self.argbegin   =3       #Excel中参数开始列索引
#self.argcount  =8        #Excel中支持的参数个数
class create_excel:
    def __init__(self, sFile, dtitleindex=3, dcasebegin=4, dargbegin=3, dargcount=8):
        self.xlApp = win32com.client.Di`spatch('et.Application')   #MS:Excel  WPS:et
        try:
            self.book = self.xlApp.Workbooks.Open(sFile)
        except:
            print_error_info()
            print "打开文件失败"
            exit()
        self.file=sFile
        self.titleindex=dtitleindex
        self.casebegin=dcasebegin
        self.argbegin=dargbegin
        self.argcount=dargcount
        self.allresult=[]

        self.retCol=self.argbegin+self.argcount
        self.xmlCol=self.retCol+1
        self.resultCol=self.xmlCol+1

    def close(self):
        #self.book.Close(SaveChanges=0)
        self.book.Save()
        self.book.Close()
        #self.xlApp.Quit()
        del self.xlApp

    def read_data(self, iSheet, iRow, iCol):
        try:
            sht = self.book.Worksheets(iSheet)
            sValue=str(sht.Cells(iRow, iCol).Value)
        except:
            self.close()
            print('读取数据失败')
            exit()
        #去除'.0'
        if sValue[-2:]=='.0':
            sValue = sValue[0:-2]
        return sValue

    def write_data(self, iSheet, iRow, iCol, sData, color=OK_COLOR):
        try:
            sht = self.book.Worksheets(iSheet)
            sht.Cells(iRow, iCol).Value = sData.decode("utf-8")
            sht.Cells(iRow, iCol).Interior.Color=color
            self.book.Save()
        except:
            self.close()
            print('写入数据失败')
            exit()

    #获取用例个数
    def get_ncase(self, iSheet):
        try:
            return self.get_nrows(iSheet)-self.casebegin+1
        except:
            self.close()
            print('获取Case个数失败')
            exit()

    def get_nrows(self, iSheet):
        try:
            sht = self.book.Worksheets(iSheet)
            return sht.UsedRange.Rows.Count
        except:
            self.close()
            print('获取nrows失败')
            exit()

    def get_ncols(self, iSheet):
        try:
            sht = self.book.Worksheets(iSheet)
            return sht.UsedRange.Columns.Count
        except:
            self.close()
            print('获取ncols失败')
            exit()

    def del_testrecord(self, suiteid):
        try:
            #为提升性能特别从For循环提取出来
            nrows=self.get_nrows(suiteid)+1
            ncols=self.get_ncols(suiteid)+1
            begincol=self.argbegin+self.argcount

            #提升性能
            sht = self.book.Worksheets(suiteid)

            for row in range(self.casebegin, nrows):
                for col in range(begincol, ncols):
                    str=self.read_data(suiteid, row, col)
                    #清除实际结果[]
                    startpos = str.find('[')
                    if startpos>0:
                        str = str[0:startpos].strip()
                        self.write_data(suiteid, row, col, str, OK_COLOR)
                    else:
                        #提升性能
                        sht.Cells(row, col).Interior.Color = OK_COLOR
                #清除TestResul列中的测试结果，设置为NT
                self.write_data(suiteid, row,  self.argbegin+self.argcount+1, ' ', OK_COLOR)
                self.write_data(suiteid, row, self.resultCol, 'NT', NT_COLOR)
        except:
            self.close()
            print('清除数据失败')
            exit()

#执行调用
def HTTPInvoke(IPPort, url):
    conn = httplib.HTTPConnection(IPPort)
    conn.request("GET", url)
    rsps = conn.getresponse()
    data = rsps.read()
    conn.close()
    return data

#获取用例基本信息[Interface,argcount,[ArgNameList]]
def get_caseinfo(Data, SuiteID):
    caseinfolist=[]
    sInterface=Data.read_data(SuiteID, 1, 2)
    argcount=int(Data.read_data(SuiteID, 2, 2))

    #获取参数名存入ArgNameList
    ArgNameList=[]
    for i in range(0, argcount):
        ArgNameList.append(Data.read_data(SuiteID, Data.titleindex, Data.argbegin+i))

    caseinfolist.append(sInterface)
    caseinfolist.append(argcount)
    caseinfolist.append(ArgNameList)
    return caseinfolist

#获取输入
def get_input(Data, SuiteID, CaseID, caseinfolist):
    sArge=''
    #参数组合
    for j in range(0, caseinfolist[1]):
        if Data.read_data(SuiteID, Data.casebegin+CaseID, Data.argbegin+j) != "None":
            sArge=sArge+caseinfolist[2][j]+'='+Data.read_data(SuiteID, Data.casebegin+CaseID, Data.argbegin+j)+'&'

    #去掉结尾的&字符
    if sArge[-1:]=='&':
        sArge = sArge[0:-1]
    sInput=caseinfolist[0]+sArge    #组合全部参数
    return sInput

#结果判断
def assert_result(sReal, sExpect):
    sReal=str(sReal)
    sExpect=str(sExpect)
    if sReal==sExpect:
        return 'OK'
    else:
        return 'NG'

#将测试结果写入文件
def write_result(Data, SuiteId, CaseId, resultcol, *result):
    if len(result)>1:
        ret='OK'
        for i in range(0, len(result)):
            if result[i]=='NG':
                ret='NG'
                break
        if ret=='NG':
            Data.write_data(SuiteId, Data.casebegin+CaseId, resultcol,ret, NG_COLOR)
        else:
            Data.write_data(SuiteId, Data.casebegin+CaseId, resultcol,ret, OK_COLOR)
        Data.allresult.append(ret)
    else:
        if result[0]=='NG':
            Data.write_data(SuiteId, Data.casebegin+CaseId, resultcol,result[0], NG_COLOR)
        elif result[0]=='OK':
            Data.write_data(SuiteId, Data.casebegin+CaseId, resultcol,result[0], OK_COLOR)
        else:  #NT
            Data.write_data(SuiteId, Data.casebegin+CaseId, resultcol,result[0], NT_COLOR)
        Data.allresult.append(result[0])

    #将当前结果立即打印
    print 'case'+str(CaseId+1)+':', Data.allresult[-1]

#打印测试结果
def statisticresult(excelobj):
    allresultlist=excelobj.allresult
    count=[0, 0, 0]
    for i in range(0, len(allresultlist)):
        #print 'case'+str(i+1)+':', allresultlist[i]
        count=countflag(allresultlist[i],count[0], count[1], count[2])
    print 'Statistic result as follow:'
    print 'OK:', count[0]
    print 'NG:', count[1]
    print 'NT:', count[2]

#解析XmlString返回Dict
def get_xmlstring_dict(xml_string):
    xml = XML2Dict()
    return xml.fromstring(xml_string)

#解析XmlFile返回Dict
def get_xmlfile_dict(xml_file):
    xml = XML2Dict()
    return xml.parse(xml_file)

#去除历史数据expect[real]
def delcomment(excelobj, suiteid, iRow, iCol, str):
    startpos = str.find('[')
    if startpos>0:
        str = str[0:startpos].strip()
        excelobj.write_data(suiteid, iRow, iCol, str, OK_COLOR)
    return str

#检查每个item （非结构体）
def check_item(excelobj, suiteid, caseid,real_dict, checklist, begincol):
    ret='OK'
    for checkid in range(0, len(checklist)):
        real=real_dict[checklist[checkid]]['value']
        expect=excelobj.read_data(suiteid, excelobj.casebegin+caseid, begincol+checkid)

        #如果检查不一致测将实际结果写入expect字段，格式：expect[real]
        #将return NG
        result=assert_result(real, expect)
        if result=='NG':
            writestr=expect+'['+real+']'
            excelobj.write_data(suiteid, excelobj.casebegin+caseid, begincol+checkid, writestr, NG_COLOR)
            ret='NG'
    return ret

#检查结构体类型
def check_struct_item(excelobj, suiteid, caseid,real_struct_dict, structlist, structbegin, structcount):
    ret='OK'
    if structcount>1:  #传入的是List
        for structid in range(0, structcount):
            structdict=real_struct_dict[structid]
            temp=check_item(excelobj, suiteid, caseid,structdict, structlist, structbegin+structid*len(structlist))
            if temp=='NG':
                ret='NG'

    else: #传入的是Dict
        temp=check_item(excelobj, suiteid, caseid,real_struct_dict, structlist, structbegin)
        if temp=='NG':
            ret='NG'

    return ret

#获取异常函数及行号
def print_error_info():
    """Return the frame object for the caller's stack frame."""
    try:
        raise Exception
    except:
        f = sys.exc_info()[2].tb_frame.f_back
    print (f.f_code.co_name, f.f_lineno)

#测试结果计数器，类似Switch语句实现
def countflag(flag,ok, ng, nt):
    calculation  = {'OK':lambda:[ok+1, ng, nt],
                         'NG':lambda:[ok, ng+1, nt],
                         'NT':lambda:[ok, ng, nt+1]}
    return calculation[flag]()

#保存XML文件
def saveXmlFile(file, xmlstring):
f = open(file, 'wb')
f.write(xmlstring)
f.close()

#检查返回值result_code
def checkResult(excelobj, sheetid, caseid, real, begincol):
ret = 'OK'
exp = excelobj.read_data(sheetid, excelobj.casebegin + caseid, begincol)

#如果检查不一致测将实际结果写入expect字段，格式：exp[real]
#将return NG
result = assert_result(real, exp)
if result == 'NG':
writestr = exp + '[' + real + ']'
excelobj.write_data(sheetid, excelobj.casebegin + caseid, begincol, writestr, NG_COLOR)
ret = 'NG'
return ret

#检查xml文件
def checkXmlFile(excelobj, sheetid, caseid, file1, file2):
ret = 'OK'
if not(filecmp.cmp(file1, file2)):
ret = 'NG'
excelobj.write_data(sheetid, excelobj.casebegin + caseid, excelobj.argbegin + excelobj.argcount + 1, 'Error', NG_COLOR)
return ret
