#encoding: utf-8

import sys
reload(sys)
sys.setdefaultencoding( "utf-8" )

import pandas as pd
import numpy as np
import time
import openpyxl


class DataProcess(object):

    def __init__(self):
        self.dtype = {u'基金代码':str,u'清算日期':str,u'委托日期':str,u'账户部门':str}
        self.branch_dict = self.get_branch_dict()
        self.total_amt_list = []
        self.daily_amt_list = []

    def get_fundcode(self):

        fundcode_str = raw_input("input fundcode: \n")
        fundcode_arr = fundcode_str.split(',')
        for fundcode in fundcode_arr:
            self.process_data(fundcode)

    def process_data(self,fundcode):
        self.daily_amt_list = []
        self.total_amt_list = []
        print fundcode
        df = pd.read_excel("Data.xlsx",converters=self.dtype)
        df = df[df[u'基金代码'] == fundcode]

        for branch_no in self.branch_list_no:
            print branch_no
            df_branch = df[df[u'账户部门'] == str(branch_no)]
            total_amt  = df_branch[u'委托金额'].sum()
            self.total_amt_list.append(total_amt)

            daily_amt = df_branch[df_branch[u'委托日期'] == self.GetNowTime()][u'委托金额'].sum()
            if daily_amt == 0:
                daily_amt = df_branch[df_branch[u'委托日期'] == self.GetNowDate()][u'委托金额'].sum()
            if daily_amt == 0:
                daily_amt = df_branch[df_branch[u'委托日期'] == self.GetNowTime2()][u'委托金额'].sum()
            if daily_amt == 0:
                daily_amt = df_branch[df_branch[u'委托日期'] == self.GetNowDate2()][u'委托金额'].sum()
            self.daily_amt_list.append(daily_amt)

        #data = {'branch_no':self.branch_list_no,'branch_name':self.branch_list_name,'today':self.daily_amt_list,'total':self.total_amt_list}
        data = {'branch_no':self.branch_list_no,'today':self.daily_amt_list,'total':self.total_amt_list}

        df_final = pd.DataFrame(data)
        df_final['branch_name'] = df_final['branch_no'].map(self.branch_dict)

        df_final = df_final.sort(columns='total',ascending=False).reset_index()
        df_final.index += 1
        file_name = "%s.xlsx" % str(fundcode)
        df_final.to_excel(file_name)

    def get_branch_dict(self):
        branch_df = pd.read_excel(u'Dict.xlsx')
        branch_dict = dict(zip(branch_df.ix[:,0],branch_df.ix[:,1]))
        self.branch_list_no = branch_df.ix[:,0].tolist()
        self.branch_list_name = branch_df.ix[:,1].tolist()
        return branch_dict

    def get_branch_list(self):
        branch_df = pd.read_excel(u'Dict.xlsx')
        branch_list1 = branch_df.ix[:,0].tolist()
        return branch_list1

    def GetNowTime(self):
	    #return time.strftime("%Y/%m/%d 00:00:00",time.localtime(time.time()))
        return '2015-12-09 00:00:00'

    def GetNowTime2(self):
	    return time.strftime("%Y%m%d 00:00:00",time.localtime(time.time()))

    def GetNowDate(self):
	    return time.strftime("%Y/%m/%d",time.localtime(time.time()))

    def GetNowDate2(self):
	    return time.strftime("%Y%m%d",time.localtime(time.time()))


if __name__ == "__main__":
    process = DataProcess()
    process.get_fundcode()