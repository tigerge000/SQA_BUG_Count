#coding=utf8
import requests
import MySQLdb
'''
Create by 古月随笔
'''
class op_mysql(object):
    def __init__(self,host,user,pwd,db):
        self.host = host
        self.user = user
        self.pwd = pwd
        self.db = db

    def __mysql_connnect(self):
        if not "huala_test":
            raise(NameError,"没有设置数据库信息")
        self.conn = MySQLdb.connect(host=self.host,user=self.user,passwd=self.pwd,db=self.db,charset="utf8",port=3306)

        cur = self.conn.cursor()

        if not cur:
            raise(NameError,"连接数据库失败")
        else:
            return cur

    '''
    查询
    '''
    def ExecQuery(self,sql):

        cur = self.__mysql_connnect()
        cur.execute(sql)
        reslist = cur.fetchall()
        #关闭连接
        self.conn.close()
        return reslist
    '''
    插入、删除、更新等操作
    '''
    def ExecNonQuery(self,sql):
        """
        执行非查询语句
        调用示例：
        cur = self.__mysql_connnect()
        cur.execute(sql)
        self.conn.commit()
        self.conn.close()
        """
        cur = self.__mysql_connnect()
        cur.execute(sql)
        self.conn.commit()
        self.conn.close()

if __name__ == "__main__":
    ms = op_mysql(host="192.168.200.166",user="erp",pwd="erpuser",db="vsmp")