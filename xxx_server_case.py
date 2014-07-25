# -*- coding: utf-8 -*-

#****************************************************************
# xxx_save_case.py
# Author     : kiven
# Version    : 1.1.2
# Date       : 2014-07-21
# Description: 自动化测试平台
#****************************************************************

from testframe import *
from common_lib import *

httpString='http://xxx.com/xxx_product/test/'
# os.getcwd() 获取python根目录
expectXmldir=os.getcwd()+'/TestDir/expect/'
realXmldir=os.getcwd()+'/TestDir/real/'

def run(interface_name, suiteid):
    print '【'+interface_name+'】' + ' Test Begin,please waiting...'
    global expectXmldir, realXmldir

    #根据接口名分别创建预期结果目录和实际结果目录
    expectDir=expectXmldir+interface_name
    realDir=realXmldir+interface_name
    if os.path.exists(expectDir) == 0:
        os.makedirs(expectDir)
    if os.path.exists(realDir) == 0:
        os.makedirs(realDir)

    excelobj.del_testrecord(suiteid)  #清除历史测试数据
    casecount=excelobj.get_ncase(suiteid) #获取case个数
    caseinfolist=get_caseinfo(excelobj, suiteid) #获取Case基本信息

    #遍历执行case
    for caseid in range(0, casecount):
        #检查是否执行该Case
        if excelobj.read_data(suiteid,excelobj.casebegin+caseid, 2)=='N':
            write_result(excelobj, suiteid, caseid, excelobj.resultCol, 'NT')
            continue #当前Case结束，继续执行下一个Case

        #获取测试数据
        sInput=httpString+get_input(excelobj, suiteid, caseid, caseinfolist)
        XmlString=HTTPInvoke(com_ipport, sInput)     #执行调用

        #获取返回码并比较
        result_code=et.fromstring(XmlString).find("result_code").text
        ret1=check_result(excelobj, suiteid, caseid,result_code, excelobj.retCol)

        #保存预期结果文件
        expectPath=expectDir+'/'+str(caseid+1)+'.xml'
        #saveXmlfile(expectPath, XmlString)

        #保存实际结果文件
        realPath=realDir+'/'+str(caseid+1)+'.xml'
        saveXmlfile(realPath, XmlString)

        #比较预期结果和实际结果
        ret2= check_xmlfile(excelobj, suiteid, caseid,expectPath, realPath)

        #写测试结果
        write_result(excelobj, suiteid, caseid, excelobj.resultCol, ret1, ret2)
    print '【'+interface_name+'】' + ' Test End!'
