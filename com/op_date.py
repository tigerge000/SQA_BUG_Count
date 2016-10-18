#coding = utf-8
import datetime
#a = datetime.datetime.now()
'''
Create by 古月随笔
'''
class op_date(object):
    def day_get(self,d):
        if type(d).__name__ == "str":
            d = datetime.datetime.strptime(d,'%Y-%m-%d %H:%M:%S')
        oneday = datetime.timedelta(days=1)
        day = d - oneday
        date_from = datetime.datetime(day.year, day.month, day.day, 0, 0, 0)
        date_to = datetime.datetime(day.year, day.month, day.day, 23, 59, 59)
        print '---'.join([str(date_from), str(date_to)])

    def week_get(self,d):
        if type(d).__name__ == "str":
            d = datetime.datetime.strptime(d,'%Y-%m-%d %H:%M:%S')
        dayscount = datetime.timedelta(days=d.isoweekday())
        dayto = d - dayscount
        sixdays = datetime.timedelta(days=6)
        dayfrom = dayto - sixdays
        date_from = datetime.datetime(dayfrom.year, dayfrom.month, dayfrom.day, 0, 0, 0)
        date_to = datetime.datetime(dayto.year, dayto.month, dayto.day, 23, 59, 59)
        datelist=[[str(date_from)],[str(date_to)]]
        # print '---'.join([str(date_from), str(date_to)])
        return datelist

    def multi_week_get(self,d,num):
        if type(d).__name__ == "str":
            d = datetime.datetime.strptime(d,'%Y-%m-%d %H:%M:%S')
        date_num = []
        # date_num = [date_to1]
        for i in range(num-1,0,-1):
            dayscount = datetime.timedelta(days=d.isoweekday())
            dayto = d - dayscount
            sixdays = datetime.timedelta(days=6*i)
            dayfrom = dayto - sixdays
            date_from = str(datetime.datetime(dayfrom.year, dayfrom.month, dayfrom.day, 10, 0, 0))
            date_num.append(date_from)

        dayscount1 = datetime.timedelta(days=d.isoweekday())
        onedays = datetime.timedelta(days=1)
        dayto1 = d - dayscount1 + onedays
        date_to1 = str(datetime.datetime(dayto1.year, dayto1.month, dayto1.day, 10, 0, 0))
        date_num.append(date_to1)
        return date_num

    def month_get(self,d):
        if type(d).__name__ == "str":
            d = datetime.datetime.strptime(d,'%Y-%m-%d %H:%M:%S')
        dayscount = datetime.timedelta(days=d.day)
        dayto = d - dayscount
        date_from = datetime.datetime(dayto.year, dayto.month, 1, 0, 0, 0)
        date_to = datetime.datetime(dayto.year, dayto.month, dayto.day, 23, 59, 59)
        # print '---'.join([str(date_from), str(date_to)])
        datelist=[[str(date_from)],[str(date_to)]]
        return datelist

if __name__ == "__main__":
    op_date = op_date()
    # print op_date.week_get("2016-01-09 23:00:00")[0]
    #print op_date.week_get("2016-01-3 23:00:00")[1]
    print op_date.multi_week_get("2015-12-23 23:00:00",4)
