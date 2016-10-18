#coding:utf-8
#__author__:huqingen

import sys
reload(sys)
sys.setdefaultencoding('utf-8')
sys.path.append("../")
import smtplib,mimetypes
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from com.op_date import *
op_date = op_date()

'''
Create by 古月随笔
'''
class sendreport():
    def send_report_by_smtp(self):
        dateNow1 = datetime.datetime.now()
        dateResult = op_date.week_get(dateNow1)
        start_date = dateResult[0][0][:-9]
        end_date = dateResult[1][0][:-9]

        msg = MIMEMultipart('alternative')
        #发件人
        mail_from = "xxx@163.com"
        # 收件人
        mail_to_list = ['hqe0079@xinyunlian.com']
        # 主题
        msg['Subject'] = u"SQA统计报告%s~%s"%(start_date,end_date)
        # 邮箱密码
        pwd = "xxxx"

        xPath1 = "report/CountBUGForWeekly%s--%s.xlsx"%(start_date,end_date)
        xPath2 = "report/MultiCountBUGForWeekly%s.xlsx"%(end_date)

        htmltext=MIMEText(u"SQA统计报告",'html','utf-8')
        msg.attach(htmltext)


		
        #添加邮件附件
		#filename="../yyTestCases.xls"
        ctype,encoding = mimetypes.guess_type(xPath1)
        if ctype is None or encoding is not None:
            ctype='application/octet-stream'
        maintype,subtype = ctype.split('/',1)
        #添加excel附件
        att1=MIMEImage(open(xPath1, 'rb').read(),subtype)
        att1["Content-Disposition"] = 'attachmemt;filename="CountBUGForWeekly%s--%s.xlsx"'%(start_date,end_date)
        msg.attach(att1)


        #添加邮件附件
        ctype,encoding = mimetypes.guess_type(xPath2)
        if ctype is None or encoding is not None:
            ctype='application/octet-stream'
        maintype,subtype = ctype.split('/',1)
        #添加excel附件
        att2=MIMEImage(open(xPath2, 'rb').read(),subtype)
        att2["Content-Disposition"] = 'attachmemt;filename="MultiCountBUGForWeekly%s.xlsx"'%(end_date)
        msg.attach(att2)


        #发送邮件
        smtp=smtplib.SMTP()
        smtp.connect("smtphz.qiye.163.com")
        smtp.login(mail_from,pwd)
        smtp.sendmail(mail_from,mail_to_list,msg.as_string())
        smtp.quit()
        print u'邮件方式发送测试报告成功'

if __name__=='__main__':
    sr=sendreport()
    sr.send_report_by_smtp()