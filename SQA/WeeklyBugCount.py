#coding=utf-8

from com.config import *
from com.op_date import *
import xlsxwriter
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
sys.path.append("../")
cn = config()
op_date = op_date()
'''
Create by 古月随笔
'''
class excelchartbyweek(object):
    def __init__(self,xpath):
        self.workbook = xlsxwriter.Workbook(xpath)
    '''
    创建图表图形方法--按周
    '''
    def chart_series_week(self,sheet_name,type,row_len,col_len):
        chart = self.workbook.add_chart({'type': '%s'%(type)})
        if type == "pie":
            for j in range(2,col_len-2):
                chart.add_series({
                    'name':       ['%s'%(sheet_name), 0, j],
                    'categories': ['%s'%(sheet_name), 1, 0, row_len, 0],
                    'values':     ['%s'%(sheet_name), 1, j, row_len,j],
                    'data_labels': {'percentage': 1},   #百分比显示数值
                })
        else:
            for j in range(2,col_len-2):
                chart.add_series({
                    'name':       ['%s'%(sheet_name), 0, j],
                    'categories': ['%s'%(sheet_name), 1, 0, row_len, 0],
                    'values':     ['%s'%(sheet_name), 1, j, row_len,j],#显示数据
                    'data_labels': {'value': 1},#显示数据表
                })
        #添加数据表
        chart.set_table()
        # 设置图表风格.
        chart.set_style(18)
        #设置图表大小
        chart.set_size({'width': 650, 'height': 450})
        return chart

    '''
    创建图表图形方法-按产品或项目
    '''
    def chart_series_all(self,sheet_name,type,row_len,col_len):
        chart = self.workbook.add_chart({'type': '%s'%(type)})
        if type == "pie":
            for j in range(col_len-4,col_len):
                chart.add_series({
                    'name':       ['%s'%(sheet_name), 0, j],
                    'categories': ['%s'%(sheet_name), 1, 0, row_len, 0],
                    'values':     ['%s'%(sheet_name), 1, j, row_len,j],
                    'data_labels': {'percentage': 1},   #百分比显示数值
                })
        else:
            for j in range(col_len-4,col_len):
                chart.add_series({
                    'name':       ['%s'%(sheet_name), 0, j],
                    'categories': ['%s'%(sheet_name), 1, 0, row_len, 0],
                    'values':     ['%s'%(sheet_name), 1, j, row_len,j],#显示数据
                    'data_labels': {'value': 1},#显示数据表
                })
        #添加数据表
        chart.set_table()
        # 设置图表风格.
        chart.set_style(18)
        #设置图表大小
        chart.set_size({'width': 650, 'height': 450})
        return chart
    '''
    柱形图
    哼哈BUG按周统计图
    @sheet_name: Sheet页名称
    @sql_date: 2016-01-04 00:00:00格式
    例:今天为2016-01-04 00:00:00，输入这个时间后，会自动查询2015-12-28 00:00:00---2016-01-03 23:59:59时间段内BUG
    '''
    def CountBUGAsWeeklyForHuaLa(self,sheet_name,sql_date):
        #计算开始时间和结束时间
        dateResult = op_date.week_get(sql_date)
        start_date = dateResult[0][0]
        end_date = dateResult[1][0]
        workbook = self.workbook
        worksheet = self.workbook.add_worksheet(name=sheet_name)
        bold = workbook.add_format({'bold': 1})
        # 定义数据表头列表

        # title = [u'按周统计图', u'统计日期',u'新增', u'已解决',u'已关闭',u'未解决(累计)',u'延期解决(累计)',u'已关闭(累计)',u'总BUG数']
        # buname = [u"哼哈微信端",u"哼哈商户端(android)",u"哼哈商户端(iOS)",u"哼哈后台",u"哼哈生活(产品)"]
        title = cn.bugStatusList
        buname = cn.huala_week
        #获取row长度
        row_len = len(buname)
        #获取col长度
        col_len = len(title)
        #定义数据列表
        #花啦微信端统计所有BUG

        data = []
        #添加哼哈微信用户端BUG数据
        result1 = cn.BugCountByProject(cn.hl_pjct[0],sql_date)
        data.append(result1)
        #添加哼哈商户端(android)BUG数据
        result2 = cn.BugCountByProject(cn.hl_pjct[1],sql_date)
        data.append(result2)
        #添加哼哈商户端(iOS)BUG数据
        result3 = cn.BugCountByProject(cn.hl_pjct[2],sql_date)
        data.append(result3)

        #添加哼哈运营后台BUG数据
        result4 = cn.BugCountByProject(cn.hl_pjct[3],sql_date)
        data.append(result4)

        #添加哼哈生活BUG数据
        result5 = cn.BugCountByProduct(cn.hl_pdct[0],sql_date)
        data.append(result5)

        format_title=workbook.add_format()    #定义format_title格式对象
        format_title.set_border(1)   #定义format_title对象单元格边框加粗(1像素)的格式
        format_title.set_bg_color('#cccccc')   #定义format_title对象单元格背景颜色为
                                               #'#cccccc'的格式
        format_title.set_align('center')    #定义format_title对象单元格居中对齐的格式
        format_title.set_bold()    #定义format_title对象单元格内容加粗的格式


        worksheet.write_row('A1', title, format_title)

        worksheet.write_column('A2', buname,bold)
        for i in range(2,row_len+2):
            worksheet.write_row('B%d'%(i),data[i-2])

        #创建一个图表，类型是column(柱形图)
        chart = self.chart_series_week(sheet_name,"column",row_len,col_len)

        # Add a chart title and some axis labels.
        chart.set_title ({'name': u'按周统计BUG %s--%s'%(start_date,end_date)})
        chart.set_x_axis({'name': u'BUG状态'})
        chart.set_y_axis({'name': u'BUG数'})

        # Insert the chart into the worksheet (with an offset).
        worksheet.insert_chart('A9', chart, {'x_offset': 25, 'y_offset': 10})


        #创建一个图表，类型是column(柱形图)
        chart1 = self.chart_series_all(sheet_name,"column",row_len,col_len)

        # Add a chart title and some axis labels.
        chart1.set_title ({'name': u'按产品或项目统计总BUG %s'%(end_date)})
        chart1.set_x_axis({'name': u'BUG状态'})
        chart1.set_y_axis({'name': u'BUG数'})

        # Insert the chart into the worksheet (with an offset).
        worksheet.insert_chart('L9', chart1, {'x_offset': 25, 'y_offset': 10})

    '''
    柱形图
    ERP&CRMBUG按周统计图
    @sheet_name: Sheet页名称
    @sql_date: 2016-01-04 00:00:00格式
    例:今天为2016-01-04 00:00:00，输入这个时间后，会自动查询2015-12-28 00:00:00---2016-01-03 23:59:59时间段内BUG
    '''
    def CountBUGAsWeeklyForERP(self,sheet_name,sql_date):
        #计算开始时间和结束时间
        dateResult = op_date.week_get(sql_date)
        start_date = dateResult[0][0]
        end_date = dateResult[1][0]
        workbook = self.workbook
        worksheet = self.workbook.add_worksheet(name=sheet_name)
        bold = workbook.add_format({'bold': 1})
        # 定义数据表头列表

        # title = [u'按周统计图', u'统计日期',u'新增', u'已解决',u'已关闭',u'未解决(累计)',u'延期解决(累计)',u'已关闭(累计)',u'总BUG数']
        # buname = [u"ERP2.0(产品)",u"CRM(产品)",u"VMP(产品)"]
        title = cn.bugStatusList
        buname = cn.erpyuvmp_week
        #获取row长度
        row_len = len(buname)
        #获取col长度
        col_len = len(title)
        #定义数据列表
        #ERP统计所有BUG
        data = []
        #添加ERP2.0 BUG数据
        result1 = cn.BugCountByProduct(cn.erp_pdct_list[0],sql_date)
        data.append(result1)
        #CRM BUG数据
        result2 = cn.BugCountByProduct(cn.erp_pdct_list[1],sql_date)
        data.append(result2)
        #VMP BUG数据
        result3 = cn.BugCountByProduct(cn.erp_pdct_list[2],sql_date)
        data.append(result3)

        format_title=workbook.add_format()    #定义format_title格式对象
        format_title.set_border(1)   #定义format_title对象单元格边框加粗(1像素)的格式
        format_title.set_bg_color('#cccccc')   #定义format_title对象单元格背景颜色为
                                               #'#cccccc'的格式
        format_title.set_align('center')    #定义format_title对象单元格居中对齐的格式
        format_title.set_bold()    #定义format_title对象单元格内容加粗的格式


        worksheet.write_row('A1', title, format_title)

        worksheet.write_column('A2', buname,bold)
        for i in range(2,row_len+2):
            worksheet.write_row('B%d'%(i),data[i-2])

        #创建一个图表，类型是column(柱形图)
        chart = self.chart_series_week(sheet_name,"column",row_len,col_len)
        # Add a chart title and some axis labels.
        chart.set_title ({'name': u'按周统计BUG %s--%s'%(start_date,end_date)})
        chart.set_x_axis({'name': u'BUG状态'})
        chart.set_y_axis({'name': u'BUG数'})
        # Insert the chart into the worksheet (with an offset).
        worksheet.insert_chart('A9', chart, {'x_offset': 25, 'y_offset': 10})


        #创建一个图表，类型是column(柱形图)
        chart1 = self.chart_series_all(sheet_name,"column",row_len,col_len)

        # Add a chart title and some axis labels.
        chart1.set_title ({'name': u'按产品或项目统计总BUG %s'%(end_date)})
        chart1.set_x_axis({'name': u'BUG状态'})
        chart1.set_y_axis({'name': u'BUG数'})

        # Insert the chart into the worksheet (with an offset).
        worksheet.insert_chart('L9', chart1, {'x_offset': 25, 'y_offset': 10})
    def teardown(self,xpath):
        self.workbook.close()
        print u"报表生成成功,报表所在路径:%s"%(xpath)

if __name__ == "__main__":
    #计算开始时间和结束时间
    print u"您现在正在使用BUG报表自动生成脚本！！！"
    dateValue = str(raw_input(u'请输入日期(参考格式:2016-01-06 11:11:11) :'))
    if dateValue == "":
        dateNow = datetime.datetime.now()
        print u"没有输入日期，默认按照当前日期生成报表，当前日期为:"
        print dateNow
        dateResult = op_date.week_get(dateNow)
        start_date = dateResult[0][0][:-9]
        end_date = dateResult[1][0][:-9]
        xpath1 = "../report/CountBUGForWeekly%s--%s.xlsx"%(start_date,end_date)
        bugcount = excelchartbyweek(xpath1)
        bugcount.CountBUGAsWeeklyForHuaLa(u"哼哈生活",dateNow)
        bugcount.CountBUGAsWeeklyForERP(u"ERP&VMP",dateNow)
        bugcount.teardown(xpath1)
    else:
        print u"您输入的日期为:"+ dateValue
        dateResult = op_date.week_get(dateValue)
        start_date = dateResult[0][0][:-9]
        end_date = dateResult[1][0][:-9]
        xpath1 = "../report/CountBUGForWeekly%s--%s.xlsx"%(start_date,end_date)
        bugcount = excelchartbyweek(xpath1)
        bugcount.CountBUGAsWeeklyForHuaLa(u"哼哈生活",dateValue)
        bugcount.CountBUGAsWeeklyForERP(u"ERP&VMP",dateValue)
        bugcount.teardown(xpath1)
