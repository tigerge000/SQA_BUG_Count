#coding=utf-8

from op_date import *
from op_mysql import *
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
'''
Create by 古月随笔
'''
class config(object):
    def __init__(self):
        self.op_date = op_date()
        # self.ms = op_mysql(host="127.0.0.1",user="root",pwd="",db="zentao")
        self.ms = op_mysql(host="192.168.200.60",user="root",pwd="123456",db="zentao")

    #BUG状态分类:
    bugStatusList = [u'按周统计图', u'统计日期',u'新增', u'已解决',u'已关闭',u'未解决(累计)',u'延期解决(累计)',u'已关闭(累计)',u'总BUG数']
    #xxxx产品&项目
    hengha = [[u"哼哈微信端"],[u"哼哈商户端(android)"],[u"哼哈商户端(iOS)"],[u"哼哈后台"],[u"哼哈生活(产品)"]]
    #花xxxxx产品&项目
    hengha_week = [u"哼哈微信端",u"哼哈商户端(android)",u"哼哈商户端(iOS)",u"哼哈后台",u"哼哈生活(产品)"]

    #ERP
    erpyuvmp = [[u"ERP2.0(产品)"],[u"CRM(产品)"],[u"VMP(产品)"]]
    #ERP
    erpyuvmp_week = [u"ERP2.0(产品)",u"CRM(产品)",u"VMP(产品)"]


    def BugCountByProject(self,projectNo,sql_date):
        ms = self.ms
        data = []
        projectNo = int(projectNo)
        date_result = self.op_date.week_get(sql_date)
        start_date = date_result[0][0]
        end_date = date_result[1][0]
        #查找一个星期内新增的BUG数openedDate 例如今天为2016-01-04 00:00:00，输入这个时间后，会自动查询2015-12-28 00:00:00---2016-01-03 23:59:59时间段内BUG
        AllNewBugCount_OneWeek = "select count(*) from zt_bug where project = '%d' and deleted = '0' and openedDate >= '%s' and openedDate <= '%s'"%(projectNo,start_date,end_date)
        #查找一个星期内已解决的BUG数(以最近的星期天为准，计算星期一到星期天,包含本周 解决到关闭的BUG) resolvedDate
        AllResolvedBugCount_OneWeek = "select count(*) from zt_bug where project = '%d' and deleted = '0' and `status` <> 'active' and resolution <> 'postponed' and resolvedDate >= '%s' and resolvedDate <= '%s'"%(projectNo,start_date,end_date)
        #查找所有未解决BUG数(以最近的星期天为准，计算星期一到星期天)（当前显示BUG状态为未解决的。包含当前还没被解决的、之前遗留的未解决、以及reopen的BUG（累计数据））
        AllNotResolvedBugCount = "select count(*) from zt_bug where project = '%d' and deleted = '0' and `status` =  'active' and openedDate <= '%s'"%(projectNo,end_date)
        #查找用户所有延期解决的问题
        AllPostponedBugCount = "select count(*) from zt_bug where project = '%d' and deleted = '0' and `status` <> 'closed' and resolution = 'postponed' and resolvedDate <= '%s'"%(projectNo,end_date)
        #查找 一个星期内已关闭的BUG数(以最近的星期天为准，计算星期一到星期天) closedDate
        AllClosedBugCount_OneWeek = "select count(*) from zt_bug where project  = '%d' and deleted = '0' and `status` = 'closed' and closedDate >= '%s' and closedDate <= '%s'"%(projectNo,start_date,end_date)

        #查找 已关闭BUG数(累计)
        AllClosedBugCount = "select count(*) from zt_bug where project  = '%d' and deleted = '0' and `status` = 'closed' and closedDate <= '%s'"%(projectNo,end_date)

        #查找 总BUG数
        AllBugCount = "select count(*) from zt_bug where project  = '%d' and deleted = '0' and openedDate <='%s'"%(projectNo,end_date)

        #新增
        dAllNewBugCount_OneWeek = ms.ExecQuery(AllNewBugCount_OneWeek)[0][0]
        #已解决
        dAllResolvedBugCount_OneWeek = ms.ExecQuery(AllResolvedBugCount_OneWeek)[0][0]
        #已关闭
        dAllClosedBugCount_OneWeek = ms.ExecQuery(AllClosedBugCount_OneWeek)[0][0]
        #未解决(累计数据)
        dAllNotResolvedBugCount = ms.ExecQuery(AllNotResolvedBugCount)[0][0]
        #延期解决(累计数据)
        dAllPostponedBugCount = ms.ExecQuery(AllPostponedBugCount)[0][0]
        #已关闭(累计)
        dAllClosedBugCount = ms.ExecQuery(AllClosedBugCount)[0][0]
        #总BUG数
        dAllBugCount = ms.ExecQuery(AllBugCount)[0][0]
        data = ["%s~%s"%(start_date[:-9],end_date[:-9]),dAllNewBugCount_OneWeek,dAllResolvedBugCount_OneWeek,dAllClosedBugCount_OneWeek,dAllNotResolvedBugCount,dAllPostponedBugCount,dAllClosedBugCount,dAllBugCount]
        return data


    def BugCountByProduct(self,productNo,sql_date):
        ms = self.ms
        data = []
        productNo = int(productNo)
        date_result = self.op_date.week_get(sql_date)
        start_date = date_result[0][0]
        end_date = date_result[1][0]
        #查找一个星期内新增的BUG数openedDate 例如今天为2016-01-04 00:00:00，输入这个时间后，会自动查询2015-12-28 00:00:00---2016-01-03 23:59:59时间段内BUG
        AllNewBugCount_OneWeek = "select count(*) from zt_bug where product = '%d' and deleted = '0' and openedDate >= '%s' and openedDate <= '%s'"%(productNo,start_date,end_date)
        #查找一个星期内已解决的BUG数(以最近的星期天为准，计算星期一到星期天) resolvedDate
        AllResolvedBugCount_OneWeek = "select count(*) from zt_bug where product = '%d' and deleted = '0' and `status` <> 'active' and resolution <> 'postponed' and resolvedDate >= '%s' and resolvedDate <= '%s'"%(productNo,start_date,end_date)
        #查找 一个星期内已关闭的BUG数(以最近的星期天为准，计算星期一到星期天) closedDate
        AllClosedBugCount_OneWeek = "select count(*) from zt_bug where product  = '%d' and deleted = '0' and `status` = 'closed' and closedDate >= '%s' and closedDate <= '%s'"%(productNo,start_date,end_date)

        #查找所有未解决BUG数(以最近的星期天为准，计算星期一到星期天)（当前显示BUG状态为未解决的。包含当前还没被解决的、之前遗留的未解决、以及reopen的BUG（累计数据））
        AllNotResolvedBugCount = "select count(*) from zt_bug where product = '%d' and deleted = '0' and `status` =  'active' and openedDate <= '%s'"%(productNo,end_date)
        #查找用户所有延期解决的问题
        AllPostponedBugCount = "select count(*) from zt_bug where product = '%d' and deleted = '0' and `status` <> 'closed' and resolution = 'postponed'and resolvedDate <= '%s'"%(productNo,end_date)
        #查找 已关闭BUG数(累计)
        AllClosedBugCount = "select count(*) from zt_bug where product  = '%d' and deleted = '0' and `status` = 'closed' and closedDate <= '%s'"%(productNo,end_date)

        #查找 总BUG数
        AllBugCount = "select count(*) from zt_bug where product  = '%d' and deleted = '0'and openedDate <='%s'"%(productNo,end_date)

        #新增
        dAllNewBugCount_OneWeek = ms.ExecQuery(AllNewBugCount_OneWeek)[0][0]
        #已解决
        dAllResolvedBugCount_OneWeek = ms.ExecQuery(AllResolvedBugCount_OneWeek)[0][0]
        #已关闭
        dAllClosedBugCount_OneWeek = ms.ExecQuery(AllClosedBugCount_OneWeek)[0][0]
        #未解决(累计数据)
        dAllNotResolvedBugCount = ms.ExecQuery(AllNotResolvedBugCount)[0][0]
        #延期解决(累计数据)
        dAllPostponedBugCount = ms.ExecQuery(AllPostponedBugCount)[0][0]
        #已关闭(累计)
        dAllClosedBugCount = ms.ExecQuery(AllClosedBugCount)[0][0]
        #总BUG数
        dAllBugCount = ms.ExecQuery(AllBugCount)[0][0]
        data = ["%s~%s"%(start_date[:-9],end_date[:-9]),dAllNewBugCount_OneWeek,dAllResolvedBugCount_OneWeek,dAllClosedBugCount_OneWeek,dAllNotResolvedBugCount,dAllPostponedBugCount,dAllClosedBugCount,dAllBugCount]

        return data
    """
    花啦生活:
        按project统计
        花啦生活-微信用户端：37
        花啦生活--商户端:39
        花啦生活--运营后台:38
    ERP&CRM(Product):
        按照产品编号来进行统计
        ERP2.0 : 3
        CRM:25
        VMP :7
    """
    '''
    花啦生活(project)
    '''
    #哼哈生活-微信用户端：37
    henghawx_pjct = 37
    #哼哈生活--商户端(android):39
    henghashandroid_pjct = 39
    #哼哈生活-商户端（IOS）:64
    henghashios_pjct = 64
    #哼哈生活--运营后台:38
    henghayy_pjct = 38
    #哼哈生活 ： 22(产品)
    hengha_pdct = 22

    hh_pjct = [henghawx_pjct,henghashandroid_pjct,henghashios_pjct,henghayy_pjct]
    hh_pdct = [hengha_pdct]

    '''
    ERP&CRM(Product):
    '''
    # ERP2.0 : 3
    erp_pdct = 3
    # CRM:25
    crm_pdct = 25
    # VMP :7
    vmp_pdct = 7
    erp_pdct_list = [erp_pdct,crm_pdct,vmp_pdct]

if __name__ == "__main__":
    cn = config()
    data = []
    result1 = cn.BugCountByProject(cn.henghawx_pjct,"2016-01-06 00:00:00")
    data.append(result1)
    result2 = cn.BugCountByProject(cn.henghash_pjct,"2016-01-06 00:00:00")
    data.append(result2)
    result3 = cn.BugCountByProject(cn.henghayy_pjct,"2016-01-06 00:00:00")
    data.append(result3)
    print data