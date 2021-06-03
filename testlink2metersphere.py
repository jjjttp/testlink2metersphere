import re
import os
import sys
import xlwt
import xml.dom.minidom
from ConvertfileCore.ExcelTestSuite import ExcelTestSuite
from ConvertfileCore.ExcelTestCase import ExcelTestCase
from ConvertfileCore.ExcelTestStep import ExcelTestStep
class GetXmlTestsuiteList(object):
    #功能模块列表
    testsuitelist=[]
    def GetXmlTestsuiteList(self,GetDirFile):
        DOMTree = xml.dom.minidom.parse(GetDirFile)
        for dom in DOMTree.childNodes:
            for d in dom.childNodes:
                if d.nodeName=="testsuite":
                    Exceltestsuite=ExcelTestSuite()
                    #获取第一级的功能模块名称
                    Exceltestsuite.__settestsuitename__(d.getAttribute("name"))
                    #获取第一级的功能模块描述
                    details=d.getElementsByTagName("details")
                    try:
                        detailsdatatmp1=details[0].childNodes[0].data
                        detailsdatatmp2=detailsdatatmp1[5:-5].encode("utf-8")
                        strstmp1=detailsdatatmp2.replace('<p>',"")
                        strstmp2= strstmp1.replace('</p>',"")
                        strstmp3= strstmp2.replace('<br />',"")
                        strstmp4= strstmp3.replace('\r',"")
                        strstmp5= strstmp4.replace('&ldquo;',"\"")
                        detailsdata= strstmp5.replace('&rdquo;',"\"")
                        Exceltestsuite.__settestsuitedec__(detailsdata)
                    except Exception as e:
                        detailsdata=""
                        print(u"功能模块描述",('%s' % e))
                        Exceltestsuite.__settestsuitedec__(detailsdata)
                    testcaselist=d.getElementsByTagName("testcase")
                    #测试用例个数
                    Exceltestsuite.__settestcasecount__(len(testcaselist))
                    self.testsuitelist.append(Exceltestsuite)
        return self.testsuitelist


#测试用例获取类

class GetXmlTestcaseList(object):
    #测试用例列表
    testcaselist=[]
    teststeplist=[]
    def gettestsuitechain(self,node):
        if node.parentNode != None:
            if node.parentNode.nodeName == 'testsuite':
                retval = self.gettestsuitechain(node.parentNode)
                return retval + '/' + node.parentNode.getAttribute("name")
        return ""
    def GetXmlTestcaseList(self,GetDirFile):
        DOMTree = xml.dom.minidom.parse(GetDirFile)
        #第一级的标签列表
        for dom in DOMTree.childNodes:
            #第二级的标签列表
            for d in dom.childNodes:
                #第二级的标签为testsuite
                if d.nodeName=="testsuite":
                    #一级功能模块下的
                    testcaselist=d.getElementsByTagName("testcase")
                    for testcase in testcaselist:
                        #测试用例对象初始化
                        Exceltestcase=ExcelTestCase()
                        #测试用例父级为testsuite
                        if testcase.parentNode.nodeName=="testsuite":
                            #测试用例：父级的功能模块名称
                            #Exceltestcase.__settestsuitename__(testcase.parentNode.getAttribute("name"))
                            Exceltestcase.__settestsuitename__(self.gettestsuitechain(testcase))
                            #用例编号赋值
                            testcaseid=testcase.getElementsByTagName("externalid")
                            Exceltestcase.__settestcasenob__(testcaseid[0].childNodes[0].data)
                            #测试用例标题
                            Exceltestcase.__settestcasename__(testcase.getAttribute("name"))
                            #测试用例摘要
                            summary=testcase.getElementsByTagName("summary")
                            try:
                                testcasesum=summary[0].childNodes[0].data
                                #testcasesumm=testcasesum[5:-5].encode("utf-8")
                                strstmp1= testcasesum.replace('<p>',"")
                                strstmp2= strstmp1.replace('</p>',"")
                                strstmp3= strstmp2.replace('<br />',"")
                                strstmp4= strstmp3.replace('\r',"")
                                strstmp5= strstmp4.replace('&ldquo;',"\"")
                                testcasesummary= strstmp5.replace('&rdquo;',"\"")
                                Exceltestcase.__settestcasesummary__(testcasesummary)
                            except Exception as e:
                                testcasesummary=""
                                Exceltestcase.__settestcasesummary__(testcasesummary)
                                print("测试用例摘要为空",('%s' % e))
                            #测试用例前置条件
                            preconditions=testcase.getElementsByTagName("preconditions")
                            try:
                                testcaseprecon=preconditions[0].childNodes[0].data
                                #testcasepreconp=testcaseprecon[5:-5].encode("utf-8")
                                strstmp1= testcaseprecon.replace('<p>',"")
                                strstmp2= strstmp1.replace('</p>',"")
                                strstmp3= strstmp2.replace('<br />',"")
                                strstmp4= strstmp3.replace('\r',"")
                                strstmp5= strstmp4.replace('&ldquo;',"\"")
                                testcasepreconpo= strstmp5.replace('&rdquo;',"\"")
                                Exceltestcase.__settestcaseprecon__(testcasepreconpo)
                            except Exception as e:
                                testcasepreconpo=""
                                Exceltestcase.__settestcaseprecon__(testcasepreconpo)
                            #状态
                            status=testcase.getElementsByTagName("status")
                            statustmp=status[0].childNodes[0].data
                            Exceltestcase.__settestcasestatus__(statustmp)
                            #重要性
                            importance=testcase.getElementsByTagName("importance")
                            importancetmp=importance[0].childNodes[0].data
                            Exceltestcase.__settestcasetype__(importancetmp)
                            #测试用例下的测试步骤+预期结果提取为字典列表
                            steps=testcase.getElementsByTagName("steps")
                            for step in steps:
                                #测试步骤初始化对象
                                Excelteststep=ExcelTestStep()
                                #测试步骤上级的测试用例编号
                                Excelteststep.__settestcasenob__(testcaseid[0].childNodes[0].data)
                                #测试步骤提取
                                testcasesteplist=step.getElementsByTagName("step")
                                teststepdic={}
                                teststepdictmp={}
                                for testcasestep in testcasesteplist:
                                    #测试步骤
                                    testcasestepco=testcasestep.getElementsByTagName("actions")
                                    try:
                                        testcasestepc=testcasestepco[0].childNodes[0].data
                                        #testcasestepcontmp=testcasestepc[5:-5].encode("utf-8")
                                        strstmp1= testcasestepc.replace('<p>',"")
                                        strstmp2= strstmp1.replace('</p>',"")
                                        strstmp3= strstmp2.replace('<br />',"")
                                        strstmp4= strstmp3.replace('\r',"")
                                        strstmp5= strstmp4.replace('&ldquo;',"\"")
                                        strstmp6 = strstmp5.replace('&rdquo;', "\"")
                                        strstmp7 = strstmp6.replace('&nbsp;', " ")
                                        strstmp8 = strstmp7.replace('<strong>', "")
                                        strstmp9 = strstmp8.replace('</strong>', "")
                                        testcasestepcon= re.sub('<!--.*?-->', "", strstmp9)
                                        Excelteststep.__setteststepcon__(testcasestepcon)
                                    except Exception as e:
                                        print("error: {}".format(e))
                                        #print(testcasestepco[0])
                                        testcasestepcon=""
                                        Excelteststep.__setteststepcon__(testcasestepcon)
                                    #预期结果
                                    expectedresultstmp=testcasestep.getElementsByTagName("expectedresults")
                                    try:
                                        expectedresultstmp1=expectedresultstmp[0].childNodes[0].data
                                        #expectedresultstmp2=expectedresultstmp1[5:-5].encode("utf-8")
                                        strstmp1= expectedresultstmp1.replace('<p>',"")
                                        strstmp2= strstmp1.replace('</p>',"")
                                        strstmp3= strstmp2.replace('<br />',"")
                                        strstmp4= strstmp3.replace('\r',"")
                                        strstmp5= strstmp4.replace('&ldquo;',"\"")
                                        strstmp6 = strstmp5.replace('&rdquo;', "\"")
                                        strstmp7 = strstmp6.replace('<strong>', "")
                                        strstmp8 = strstmp7.replace('</strong>', "")
                                        strstmp9 = strstmp8.replace('<em>', "")
                                        strstmp10 = strstmp9.replace('</em>', "")
                                        expectedresult= strstmp10.replace('&nbsp;'," ")
                                        Excelteststep.__setteststepexpect__(expectedresult)
                                    except Exception as e:
                                        expectedresult=""
                                        Excelteststep.__setteststepexpect__(expectedresult)
                                    dic=teststepdictmp.fromkeys([testcasestepcon],expectedresult)
                                    teststepdic.update(dic)
                                Excelteststep.__setteststepnum__(len(steps))
                                Exceltestcase.__setteststepdic__(teststepdic)
                            self.teststeplist.append(Excelteststep)
                            self.testcaselist.append(Exceltestcase)
        return self.testcaselist,self.teststeplist
# class write2xls(object):
#     def __init__(self):
#         # 创建workbook和sheet对象
#         workbook = xlwt.Workbook()  # 注意Workbook的开头W要大写
#         sheet1 = workbook.add_sheet('sheet1', cell_overwrite_ok=True)
#         return sheet1
def replace_L(match):
    return match.group(0).replace(match.group(0), "{}、".format(match.group(1)))
if __name__ == '__main__':
    xml_file_path = "testlink.xml"
    xls_file_path = ""
    if len(sys.argv) == 2:
        #sys.argv.append(xml_file_path)  # 第一个参数
        xml_file_path = sys.argv[1]
        xls_file_path = os.path.splitext(xml_file_path)[0] + "ms.xls"
    tc = GetXmlTestcaseList()
    tc_list,ts_list = tc.GetXmlTestcaseList(xml_file_path)

    # 创建workbook和sheet对象
    workbook = xlwt.Workbook()  # 注意Workbook的开头W要大写
    sheet1 = workbook.add_sheet('sheet1', cell_overwrite_ok=True)
    #Add sheet header:
    sheet1.write(0, 0,"用例名称")
    sheet1.write(0, 1, "所属模块")
    sheet1.write(0, 2, "维护人")
    sheet1.write(0, 3, "用例等级")
    sheet1.write(0, 4, "标签")
    sheet1.write(0, 5, "前置条件")
    sheet1.write(0, 6, "备注")
    sheet1.write(0, 7, "步骤描述")
    sheet1.write(0, 8, "预期结果")
    sheet1.write(0, 9, "编辑模式")

    for row,tc_item in enumerate(tc_list,1):
        sheet1.write(row, 0, tc_item.testcasename)
        sheet1.write(row, 1, tc_item.testsuitename)
        sheet1.write(row, 4, "testlink_import")
        sheet1.write(row, 2, "admin")
        sheet1.write(row, 3, "P1")
        sheet1.write(row, 5, tc_item.testcaseprecon)
        sheet1.write(row, 6, tc_item.testcasesummary)
        teststep2write=[]
        teststepexpect2write=[]
        for tstep in ts_list:
            if tstep.testcasenob == tc_item.testcasenob:
                for index,step in enumerate(tstep.teststepcon,1):
                    #将步骤中带有1.等序号替换成1、以免引起误识别：
                    step = re.sub("^([0-9]+)\.", replace_L, step, flags=re.MULTILINE)
                    teststep2write.append("{}. {}".format(index,step))
                for index,stepexpect in enumerate(tstep.teststepexpect,1):
                    stepexpect = stepexpect if stepexpect else '\n'
                    # 将步骤中带有1.等序号替换成1、以免引起误识别：
                    stepexpect = re.sub("^([0-9]+)\.", replace_L, stepexpect, flags=re.MULTILINE)
                    teststepexpect2write.append("{}. {}".format(index, stepexpect))
                sheet1.write(row, 7, teststep2write)
                sheet1.write(row, 8, teststepexpect2write)
    workbook.save(xls_file_path)
    print("转换用例数量：{}".format(len(tc_list)))

    # cl = convertxml2xls(xml_file)