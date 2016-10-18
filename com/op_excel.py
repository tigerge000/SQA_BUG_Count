#coding=utf-8
import xlrd
import sys
from xlutils.copy import copy
sys.path.append("../")

'''
excel操作方法
Create by 古月随笔
'''
class op_excel(object):
    '''
    从excel中读取要查找的商品名称，第一列默认为0
    '''
    def Read_Good_Name(self,xpath,col_index=None,sheet_index=None):
        sheet_index = int(sheet_index)
        col_index = int(col_index)
        #打开xls格式文件，并保存之前数据的格式
        rb = xlrd.open_workbook(xpath,formatting_info=True)
        #获取当前sheet页
        r_sheet = rb.sheet_by_index(sheet_index)
        #获取总行数
        table_row_nums = r_sheet.nrows
        list = []
        #进行格式转换
        for i in range(1,table_row_nums):
            #按列读取行值
            cvalue = r_sheet.cell(i,col_index).value
            if type(cvalue).__name__ == 'unicode':
                cvalue = cvalue.encode('utf-8')
            elif type(cvalue).__name__ == 'float':
                cvalue = str(int(cvalue))
            #保存到list中
            list.append(cvalue)
        return list

    def Write_Data(self,xpath,row_index =None):
        row_index = int(row_index)
        #打开xls格式文件，并保存之前数据的格式
        rb = xlrd.open_workbook(xpath,formatting_info=True)
        #获取当前sheet页
        r_sheet = rb.sheet_by_index(0)
        #拷贝变量
        wb = copy(rb)
        #根据wb获取对应是sheet
        w_sheet = wb.get_sheet(0)
        row_data = "yes"
        w_sheet.write(row_index,1,row_data)
        wb.save(xpath)

       #测试通过
    def write(self,xpath,sheet_index,iRow,iCol,sData):
        sheet_index = int(sheet_index)
        rb = xlrd.open_workbook(xpath,formatting_info=True)
        wb = copy(rb)
        w_sheet = wb.get_sheet(sheet_index)
        w_sheet.write(iRow,iCol,sData.decode("utf-8"))
        wb.save(xpath)

    def write_excel(xpath,sheetindex,iRow,iCol,sData):
        sheetindex = int(sheetindex)
        #打开xls格式文件，并保存之前数据的格式
        rb = xlrd.open_workbook(xpath,formatting_info=True)
        wb = copy(rb)
        w_sheet = wb.get_sheet(sheetindex)
        w_sheet.write(iRow,iCol,sData.decode("utf-8"))
        wb.save(xpath)


    def readExcel(self,xpath,sheetname):
        #打开xls格式文件，并保存之前数据的格式
        rb = xlrd.open_workbook(xpath,formatting_info=True)
        r_sheet = rb.sheet_by_name(sheetname)
        #获取总行数
        nrows = r_sheet.nrows
        #获取总列数
        ncols = r_sheet.ncols
        print "rows: %d and cols:%d"%(nrows,ncols)
        cases = []
        for i in range(1,nrows):
            cases.append(r_sheet.row_values(i))
        print cases
        return cases
    """
    回写测试结果
    """
    def writeExcel(self,xpath,sheetname,lineNum,ifPassed,result):
        #打开xls格式文件，并保存之前数据的格式
        rb = xlrd.open_workbook(xpath,formatting_info=True)
        #获取当前sheet页
        r_sheet = rb.sheet_by_name(sheetname)
        #拷贝变量
        wb = copy(rb)
        #根据wb获取对应是sheet
        w_sheet = wb.get_sheet(0)
        #判断若该用例已经执行失败，则不能回写测试结果
        if r_sheet.cell(lineNum,2) == "N":
            print "结果已存在"
        else:
            w_sheet.write(lineNum,2,ifPassed)
        #判断若该用例若无测试结果，则回写
        if ifPassed == "N":
            w_sheet.write(lineNum,3,result)
        else:
            print u"结果信息不符合规范"
        wb.save(xpath)


if __name__ == "__main__":
    l_list = op_excel().readExcel("../yyTestCases.xls",u"登录")
    print l_list[0][1]


