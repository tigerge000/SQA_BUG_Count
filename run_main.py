#coding=utf-8
__author__ = 'huqingen'
#date: 2016-01-18

from SQA.WeeklyBugCount import *
from SQA.MultiWeekBugCount import *
from com.sendReport import *
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

'''
Create by 古月随笔
'''
if __name__ == "__main__":

    sendEmail = sendreport()
    print u"每周一早上5点发送SQA统计报告"

    print u"您现在正在使用BUG报表自动生成脚本！！！"

    print u"现在正在生成周报(CountBUGForWeekly)"

    dateNow = datetime.datetime.now()
    # print u"没有输入日期，默认按照当前日期生成报表，当前日期为:"
    print dateNow
    dateResult = op_date.week_get(dateNow)
    start_date = dateResult[0][0][:-9]
    end_date = dateResult[1][0][:-9]
    xpath = "report/CountBUGForWeekly%s--%s.xlsx"%(start_date,end_date)
    bugcount = excelchartbyweek(xpath)
    bugcount.CountBUGAsWeeklyForHuaLa(u"花啦生活",dateNow)
    bugcount.CountBUGAsWeeklyForZhiFu(u"会员与支付",dateNow)
    bugcount.CountBUGAsWeeklyForJuDian(u"聚店&司机",dateNow)
    bugcount.CountBUGAsWeeklyForERP(u"ERP&VMP",dateNow)
    bugcount.CountBUGAsWeeklyForXYT(u"新云团",dateNow)
    bugcount.teardown(xpath)

    weekNo = 4
    print u"现在正在生成多个星期周报(MultiCountBUGForWeekly)"
    xpath1 = "report/MultiCountBUGForWeekly%s.xlsx"%(end_date)
    bugcount1 = excelchartbyMultiweek(xpath1)
    bugcount1.MultiCountBUGAsWeeklyForHuaLa(u"花啦生活",dateNow,weekNo)
    bugcount1.MultiCountBUGAsWeeklyForZhiFu(u"会员与支付",dateNow,weekNo)
    bugcount1.MultiCountBUGAsWeeklyForJuDian(u"聚店&司机",dateNow,weekNo)
    bugcount1.MultiCountBUGAsWeeklyForERP(u"ERP&VMP",dateNow,weekNo)
    bugcount1.MultiCountBUGAsWeeklyForXYT(u"新云团",dateNow,weekNo)
    bugcount1.teardown(xpath1)

    sendEmail.send_report_by_smtp()
