#coding=utf-8

from com.op_mysql import *
from com.config import *
from com.op_date import *
import xlsxwriter
import random

ms = op_mysql(host="192.168.200.60",user="root",pwd="123456",db="zentao")
cn = config()
op_date = op_date()

'''
Create by 古月随笔
'''
class excelchart(object):
    def __init__(self,xpath):
        self.workbook = xlsxwriter.Workbook(xpath)

    def get_num(self):
        return random.randrange(0, 201, 2)
    #定义图表数据系列函数

    # def chart_series(self,chart,cur_row,sheet_name,type):
    #     #chart =self.workbook.add_chart({'type': type})
    #     chart.add_series({
    #         'categories': '=%s!$B$1:$D$1'%(sheet_name),
    #         'values':     '=%s!$B$'%(sheet_name)+cur_row+':$H$'+cur_row,
    #         'line':       {'color': 'black'},
    #         'name': '=%s!$A$'%(sheet_name)+cur_row,
    #     })
    '''
    折线图
    BUG趋势图
    '''
    def CountBUGAsALLForhuala(self,sheet_name):
        workbook = self.workbook
        #worksheet = self.worksheetCountBUGAsALL
        worksheet = self.workbook.add_worksheet(name=sheet_name)
        bold = workbook.add_format({'bold': 1})
        # 定义数据表头列表
        title = [u'BUG趋势图', u'BUG总数', u'未关闭',u'关闭']
        buname = [u"花啦微信端",u"花啦商户端",u"花啦运营后台"]
        #定义数据列表
        #花啦微信端统计所有BUG
        hlwxAllBugCount = ms.ExecQuery(cn.hlwxAllBugCount)[0][0] #所有
        hlwxAllNotClosedBugCount = ms.ExecQuery(cn.hlwxAllNotClosedBugCount)[0][0]#未关闭
        hlwxAllClosedBugCount = ms.ExecQuery(cn.hlwxAllClosedBugCount)[0][0] #已关闭

        #花啦商户端
        hlshAllBugCount = ms.ExecQuery(cn.hlshAllBugCount)[0][0] #所有
        hlshAllNotClosedBugCount = ms.ExecQuery(cn.hlshAllNotClosedBugCount)[0][0]#未关闭
        hlshAllClosedBugCount = ms.ExecQuery(cn.hlshAllClosedBugCount)[0][0] #已关闭

        #花啦运营后台
        hlyyAllBugCount = ms.ExecQuery(cn.hlyyAllBugCount)[0][0] #所有
        hlyyAllNotClosedBugCount = ms.ExecQuery(cn.hlyyAllNotClosedBugCount)[0][0]#未关闭
        hlyyAllClosedBugCount = ms.ExecQuery(cn.hlyyAllClosedBugCount)[0][0] #已关闭
        data = [[hlwxAllBugCount,hlwxAllNotClosedBugCount,hlwxAllClosedBugCount],[hlshAllBugCount,hlshAllNotClosedBugCount,hlshAllClosedBugCount],
                [hlyyAllBugCount,hlyyAllNotClosedBugCount,hlyyAllClosedBugCount]]
        # for i in range(3):
        #     tmp = []
        #     for j in range(3):
        #         tmp.append(self.get_num())
        #     data.append(tmp)
        format_title=workbook.add_format()    #定义format_title格式对象
        format_title.set_border(1)   #定义format_title对象单元格边框加粗(1像素)的格式
        format_title.set_bg_color('#cccccc')   #定义format_title对象单元格背景颜色为
                                               #'#cccccc'的格式
        format_title.set_align('center')    #定义format_title对象单元格居中对齐的格式
        format_title.set_bold()    #定义format_title对象单元格内容加粗的格式


        worksheet.write_row('A1', title, format_title)
        worksheet.write_column('A2', buname,bold)
        worksheet.write_row('B2',data[0])
        worksheet.write_row('B3',data[1])
        worksheet.write_row('B4',data[2])
        #创建一个图表，类型是line(折线图)
        chart1 = workbook.add_chart({'type': 'column'})
        # for row in range(2,5):
        #     self.chart_series(chart1,row,sheet_name,"line")
        #
        # 配置series,这个和前面wordsheet是有关系的。
        chart1.add_series({
            'name':       '=%s!$B$1'%(sheet_name),
            'categories': '=%s!$A$2:$A$4'%(sheet_name),
            'values':     '=%s!$B$2:$B$4'%(sheet_name),
            'data_labels': {'value': 1},
        })
        chart1.add_series({
            'name':       '=%s!$C$1'%(sheet_name),
            'categories': '=%s!$A$2:$A$4'%(sheet_name),
            'values':     '=%s!$C$2:$C$4'%(sheet_name),
            'data_labels': {'value': 1},
        })
        chart1.add_series({
            'name':       '=%s!$D$1'%(sheet_name),
            'categories': '=%s!$A$2:$A$4'%(sheet_name),
            'values':     '=%s!$D$2:$D$4'%(sheet_name),
            'data_labels': {'value': 1},
        })

        # Add a chart title and some axis labels.
        chart1.set_title ({'name': u'BUG趋势图'})
        chart1.set_x_axis({'name': u'项目'})
        chart1.set_y_axis({'name': u'BUG数'})

        # Set an Excel chart style.
        chart1.set_style(18)

        # Insert the chart into the worksheet (with an offset).
        worksheet.insert_chart('F2', chart1, {'x_offset': 25, 'y_offset': 10})
        #workbook.close()
    '''
    柱形图
    按BUG所属模块统计图
    '''
    def CountBUGAsModule(self,sheet_name):
        workbook = self.workbook
        worksheet = self.workbook.add_worksheet(name=sheet_name)
        bold = workbook.add_format({'bold': 1})
        # 定义数据表头列表
        title = [u'按模块统计图', u'严重', u'一般',u'建议']
        buname = [u"模块一",u"模块二",u"模块三",u"模块四"]
        #定义数据列表
        data = []
        for i in range(4):
            tmp = []
            for j in range(3):
                tmp.append(self.get_num())
            data.append(tmp)
        format_title=workbook.add_format()    #定义format_title格式对象
        format_title.set_border(1)   #定义format_title对象单元格边框加粗(1像素)的格式
        format_title.set_bg_color('#cccccc')   #定义format_title对象单元格背景颜色为
                                               #'#cccccc'的格式
        format_title.set_align('center')    #定义format_title对象单元格居中对齐的格式
        format_title.set_bold()    #定义format_title对象单元格内容加粗的格式


        worksheet.write_row('A1', title, format_title)
        worksheet.write_column('A2', buname,bold)
        worksheet.write_row('B2',data[0])
        worksheet.write_row('B3',data[1])
        worksheet.write_row('B4',data[2])
        worksheet.write_row('B5',data[3])
        #创建一个图表，类型是column(柱形图)
        chart1 = workbook.add_chart({'type': 'column'})
        #创建一个图表，类型是line(折线图)
        # 配置series,这个和前面wordsheet是有关系的。
        chart1.add_series({
            'name':       '=%s!$B$1'%(sheet_name),
            'categories': '=%s!$A$2:$A$5'%(sheet_name),
            'values':     '=%s!$B$2:$B$5'%(sheet_name),
            'data_labels': {'value': 1},
        })
        chart1.add_series({
            'name':       '=%s!$C$1'%(sheet_name),
            'categories': '=%s!$A$2:$A$5'%(sheet_name),
            'values':     '=%s!$C$2:$C$5'%(sheet_name),
            'data_labels': {'value': 1},
        })
        chart1.add_series({
            'name':       '=%s!$D$1'%(sheet_name),
            'categories': '=%s!$A$2:$A$5'%(sheet_name),
            'values':     '=%s!$D$2:$D$5'%(sheet_name),
            'data_labels': {'value': 1},
        })

        # Add a chart title and some axis labels.
        chart1.set_title ({'name': u'按模块统计BUG'})
        chart1.set_x_axis({'name': u'严重等级'})
        chart1.set_y_axis({'name': u'BUG数'})

        # 设置图表风格.
        chart1.set_style(18)

        # Insert the chart into the worksheet (with an offset).
        worksheet.insert_chart('F2', chart1, {'x_offset': 25, 'y_offset': 10})
    '''
    饼图
    按BUG严重级别统计图
    '''
    def CountBUGAsSevereLevel(self,sheet_name):
        workbook = self.workbook
        worksheet = self.workbook.add_worksheet(name=sheet_name)
        bold = workbook.add_format({'bold': 1})
        # 定义数据表头列表
        title = ['', u'BUG严重级别统计']
        buname = [u"致命",u"严重",u"一般",u"建议"]

        #定义数据列表
        data = []
        for i in range(4):
            tmp = []
            tmp.append(self.get_num())
            data.append(tmp)
        format_title=workbook.add_format()    #定义format_title格式对象
        format_title.set_border(1)   #定义format_title对象单元格边框加粗(1像素)的格式
        format_title.set_bg_color('#cccccc')   #定义format_title对象单元格背景颜色为
                                               #'#cccccc'的格式
        format_title.set_align('center')    #定义format_title对象单元格居中对齐的格式
        format_title.set_bold()    #定义format_title对象单元格内容加粗的格式


        worksheet.write_row('A1', title, format_title)
        worksheet.write_column('A2', buname,bold)
        worksheet.write_row('B2',data[0])
        worksheet.write_row('B3',data[1])
        worksheet.write_row('B4',data[2])
        worksheet.write_row('B5',data[3])
        #创建一个图表，类型是pie(饼图)
        chart1 = workbook.add_chart({'type': 'pie'})
        # 配置series,这个和前面wordsheet是有关系的。
        chart1.add_series({
            'name':       '=%s!$B$1'%(sheet_name),
            'categories': '=%s!$A$2:$A$5'%(sheet_name),
            'values':     '=%s!$B$2:$B$5'%(sheet_name),
            'data_labels': {'percentage': 1},   #百分比显示数值
        })
        # 设置图表风格.
        chart1.set_style(18)
        # Insert the chart into the worksheet (with an offset).
        worksheet.insert_chart('F2', chart1, {'x_offset': 25, 'y_offset': 10})

    '''
    饼图
    按BUG类型统计图
    '''
    def CountBUGAsType(self,sheet_name):
        workbook = self.workbook
        worksheet = self.workbook.add_worksheet(name=sheet_name)
        bold = workbook.add_format({'bold': 1})
        # 定义数据表头列表
        title = ['', u'BUG类型统计']
        buname = ["UI","UE","FC","CK","IF"]

        #定义数据列表
        data = []
        for i in range(5):
            tmp = []
            tmp.append(self.get_num())
            data.append(tmp)
        format_title=workbook.add_format()    #定义format_title格式对象
        format_title.set_border(1)   #定义format_title对象单元格边框加粗(1像素)的格式
        format_title.set_bg_color('#cccccc')   #定义format_title对象单元格背景颜色为
                                               #'#cccccc'的格式
        format_title.set_align('center')    #定义format_title对象单元格居中对齐的格式
        format_title.set_bold()    #定义format_title对象单元格内容加粗的格式


        worksheet.write_row('A1', title, format_title)
        worksheet.write_column('A2', buname,bold)
        worksheet.write_row('B2',data[0])
        worksheet.write_row('B3',data[1])
        worksheet.write_row('B4',data[2])
        worksheet.write_row('B5',data[3])
        worksheet.write_row('B6',data[4])
        #创建一个图表，类型是pie(饼图)
        chart1 = workbook.add_chart({'type': 'pie'})
        # 配置series,这个和前面wordsheet是有关系的。
        chart1.add_series({
            'name':       '=%s!$B$1'%(sheet_name),
            'categories': '=%s!$A$2:$A$6'%(sheet_name),
            'values':     '=%s!$B$2:$B$6'%(sheet_name),
            'line':       {'color': 'black'},    #线条颜色定义为black(黑色)
            'data_labels': {'percentage': 1},   #百分比显示数值
        })


        # 设置图表风格.
        chart1.set_style(18)

        # Insert the chart into the worksheet (with an offset).
        worksheet.insert_chart('F2', chart1, {'x_offset': 25, 'y_offset': 10})

    '''
    柱形图
    BUG按周统计图
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

        title = [u'按周统计图', u'新增', u'已解决',u'未解决',u'延期解决',u'已关闭']
        buname = [u"花啦微信端",u"花啦商户端",u"花啦后台"]

        #定义数据列表
        #花啦微信端统计所有BUG

        #查找花啦生活--微信用户 一个星期内新增的BUG数openedDate
        #新增
        hlwxAllNewBugCount_OneWeek = ms.ExecQuery(cn.hlwxAllNewBugCount_OneWeek%(start_date,end_date))[0][0]
        #已解决
        hlwxAllResolvedBugCount_OneWeek = ms.ExecQuery(cn.hlwxAllResolvedBugCount_OneWeek%(start_date,end_date))[0][0]
        #未解决
        hlwxAllNotResolvedBugCount = ms.ExecQuery(cn.hlwxAllNotResolvedBugCount)[0][0]

        #延期解决
        hlwxAllPostponedBugCount = ms.ExecQuery(cn.hlwxAllPostponedBugCount)[0][0]
        #已关闭
        hlwxAllClosedBugCount_OneWeek = ms.ExecQuery(cn.hlwxAllClosedBugCount_OneWeek%(start_date,end_date))[0][0]


        #花啦商户端
        #新增
        hlshAllNewBugCount_OneWeek = ms.ExecQuery(cn.hlshAllNewBugCount_OneWeek%(start_date,end_date))[0][0]
        #已解决
        hlshAllResolvedBugCount_OneWeek = ms.ExecQuery(cn.hlshAllResolvedBugCount_OneWeek%(start_date,end_date))[0][0]
        #未解决
        hlshAllNotResolvedBugCount = ms.ExecQuery(cn.hlshAllNotResolvedBugCount)[0][0]
        #延期解决
        hlshAllPostponedBugCount = ms.ExecQuery(cn.hlshAllPostponedBugCount)[0][0]
        #已关闭
        hlshAllClosedBugCount_OneWeek = ms.ExecQuery(cn.hlshAllClosedBugCount_OneWeek%(start_date,end_date))[0][0]



        # #花啦运营后台
        #新增
        hlyyAllNewBugCount_OneWeek = ms.ExecQuery(cn.hlyyAllNewBugCount_OneWeek%(start_date,end_date))[0][0]
        #已解决
        hlyyAllResolvedBugCount_OneWeek = ms.ExecQuery(cn.hlyyAllResolvedBugCount_OneWeek%(start_date,end_date))[0][0]
        #未解决
        hlyyAllNotResolvedBugCount = ms.ExecQuery(cn.hlyyAllNotResolvedBugCount)[0][0]
        #延期解决
        hlyyAllPostponedBugCount = ms.ExecQuery(cn.hlyyAllPostponedBugCount)[0][0]
        #已关闭
        hlyyAllClosedBugCount_OneWeek = ms.ExecQuery(cn.hlyyAllClosedBugCount_OneWeek%(start_date,end_date))[0][0]

        data = [[hlwxAllNewBugCount_OneWeek,hlwxAllResolvedBugCount_OneWeek,hlwxAllNotResolvedBugCount,hlwxAllPostponedBugCount,hlwxAllClosedBugCount_OneWeek],
                [hlshAllNewBugCount_OneWeek,hlshAllResolvedBugCount_OneWeek,hlshAllNotResolvedBugCount,hlshAllPostponedBugCount,hlshAllClosedBugCount_OneWeek],
                [hlyyAllNewBugCount_OneWeek,hlyyAllResolvedBugCount_OneWeek,hlyyAllNotResolvedBugCount,hlyyAllPostponedBugCount,hlyyAllClosedBugCount_OneWeek]]


        format_title=workbook.add_format()    #定义format_title格式对象
        format_title.set_border(1)   #定义format_title对象单元格边框加粗(1像素)的格式
        format_title.set_bg_color('#cccccc')   #定义format_title对象单元格背景颜色为
                                               #'#cccccc'的格式
        format_title.set_align('center')    #定义format_title对象单元格居中对齐的格式
        format_title.set_bold()    #定义format_title对象单元格内容加粗的格式


        worksheet.write_row('A1', title, format_title)
        worksheet.write_column('A2', buname,bold)
        worksheet.write_row('B2',data[0])
        worksheet.write_row('B3',data[1])
        worksheet.write_row('B4',data[2])

        #创建一个图表，类型是column(柱形图)
        chart1 = workbook.add_chart({'type': 'column'})
        #创建一个图表，类型是line(折线图)
        # 配置series,这个和前面wordsheet是有关系的。
        chart1.add_series({
            'name':       '=%s!$B$1'%(sheet_name),
            'categories': '=%s!$A$2:$A$4'%(sheet_name),
            'values':     '=%s!$B$2:$B$4'%(sheet_name),
            'data_labels': {'value': 1},
        })
        chart1.add_series({
            'name':       '=%s!$C$1'%(sheet_name),
            'categories': '=%s!$A$2:$A$4'%(sheet_name),
            'values':     '=%s!$C$2:$C$4'%(sheet_name),
            'data_labels': {'value': 1},
        })
        chart1.add_series({
            'name':       '=%s!$D$1'%(sheet_name),
            'categories': '=%s!$A$2:$A$4'%(sheet_name),
            'values':     '=%s!$D$2:$D$4'%(sheet_name),
            'data_labels': {'value': 1},
        })
        chart1.add_series({
            'name':       '=%s!$E$1'%(sheet_name),
            'categories': '=%s!$A$2:$A$4'%(sheet_name),
            'values':     '=%s!$E$2:$E$4'%(sheet_name),
            'data_labels': {'value': 1},
        })
        chart1.add_series({
            'name':       '=%s!$F$1'%(sheet_name),
            'categories': '=%s!$A$2:$A$4'%(sheet_name),
            'values':     '=%s!$F$2:$F$4'%(sheet_name),
            'data_labels': {'value': 1},
        })
        # Add a chart title and some axis labels.
        chart1.set_title ({'name': u'按周统计BUG %s--%s'%(start_date,end_date)})
        chart1.set_x_axis({'name': u'BUG状态'})
        chart1.set_y_axis({'name': u'BUG数'})
        #添加数据表
        chart1.set_table()
        # 设置图表风格.
        chart1.set_style(18)
        #设置图表大小
        chart1.set_size({'width': 650, 'height': 450})

        # Insert the chart into the worksheet (with an offset).
        worksheet.insert_chart('H2', chart1, {'x_offset': 25, 'y_offset': 10})
    def teardown(self):
        self.workbook.close()


if __name__ == "__main__":
    xpath = "CountBUGForHuala.xlsx"
    bugcount = excelchart(xpath)
    bugcount.CountBUGAsALLForhuala("CountBUGAsALL")
    bugcount.CountBUGAsModule("CountBUGAsModule")
    bugcount.CountBUGAsSevereLevel("CountBUGAsSevereLevel")
    bugcount.CountBUGAsType("CountBUGAsType")
    bugcount.CountBUGAsWeeklyForHuaLa("CountBUGAsWeeklyForHuaLa","2016-01-06 00:11:11")
    bugcount.teardown()
