import os
import unittest
from shutil import copy

import xlrd
import xlwt as xlwt
import getPath# 自己定义的内部类，该类返回项目的绝对路径
#调用读Excel的第三方库xlrd
from xlrd import open_workbook
# 拿到该项目所在的绝对路径
path = getPath.get_Path()

class readExcel():
    def excel_data_list(self, xls_name, sheetname):
        data_list = []
        xlsPath = os.path.join(path, "testFile", 'case', xls_name)
        wb = open_workbook(xlsPath)  # 打开excel
        sh = wb.sheet_by_name(sheetname)  # 定位工作表
        header = sh.row_values(0)  # 获取标题行的数据
        for i in range(1, sh.nrows):  # 跳过标题行，从第二行开始获取数据
            col_datas = dict(zip(sh.row_values(0), sh.row_values(i)))  # 将标题和每一行的数据，组装成字典
            data_list.append(col_datas)  # 将字典添加到列表中 ，列表嵌套字典，相当于每个字典的元素都是一个列表（也就是一行数据）
        return data_list

    def get_test_data(self, data_list, case_id):
        '''
        :param data_list: 工作表的所有行数据
        :param case_id: 用例id，用来判断执行哪几条case。如果id=all ，那就执行所有用例；否则，执行列表参数中指定的用例
        :return:  返回最终要执行的测试用例
        '''
        if case_id == 'all':
            final_data = data_list
        else:
            final_data = []
            for item in data_list:
                if item['id'] in case_id:
                    final_data.append(item)
        return final_data
    def get_xls(self,xls_name, sheet_name):# xls_name填写用例的Excel名称 sheet_name该Excel的sheet名称
        cls = []
        # 获取用例文件路径
        xlsPath = os.path.join(path, "testFile", 'case', xls_name)
        file = open_workbook(xlsPath)# 打开用例Excel
        sheet = file.sheet_by_name(sheet_name)#获得打开Excel的sheet
        # 获取这个sheet内容行数
        nrows = sheet.nrows
        for i in range(nrows):#根据行数做循环
            if sheet.row_values(i)[0] != u'username':#如果这个Excel的这个sheet的第i行的第一列不等于case_name那么我们把这行的数据添加到cls[]
                cls.append(sheet.row_values(i))
        return cls

    def write_excel_xls_append(self,path, value):

        index = len(value)  # 获取需要写入数据的行数
        workbook = xlrd.open_workbook(path)  # 打开工作簿
        sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
        worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
        rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
        new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
        new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
        for i in range(0, index):
            for j in range(0, len(value[i])):
                new_worksheet.write(i + rows_old, j, value[i][j])  # 追加写入数据，注意是从i+rows_old行开始写入
        new_workbook.save(path)  # 保存工作簿
        print("xls格式表格【追加】写入数据成功！")


    def writeExcel(self,xls_name,sheet_value):
        index = len(sheet_value)
        workbook = xlrd.open_workbook(xls_name)  # 打开工作簿
        sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
        worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
        rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
        new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
        new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
        for i in range(0, index):
            for j in range(0, len(sheet_value[i])):
                new_worksheet.write(i + rows_old, j, sheet_value[i][j])  # 追加写入数据，注意是从i+rows_old行开始写入
        new_workbook.save(path)  # 保存工作簿
        print("xls格式表格【追加】写入数据成功！")
        # if self.sheet_name is not None:
        #     workbook = xlrd.open_workbook(xls_name)
        #     sheets = workbook.sheet_names()
        #     wb=xlwt.Workbook()
        #     sheet=wb.add_sheet(sheet_name)
        #     for i in range(len(sheet_value)):
        #         for j in range(len(sheet_value[i])):
        #             sheet.write(i,j,sheet_value[i][j])
        #
        #     wb.save(self.EXCEL_PATH)
        #     print("write date success!")
        #     return True
        # else:
        #     return False

if __name__ == '__main__':#我们执行该文件测试一下是否可以正确获取Excel中的值

    test_data=readExcel().excel_data_list('login.xlsx', 'login')
    print('第一种方法：',test_data)
    cls=readExcel().get_xls('login.xlsx', 'login')
    print(readExcel().get_xls('login.xlsx', 'login'))
    for i in cls:
        for j in range(len(i)):
            print(i[j])

    unittest.main()