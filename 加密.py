#coding:gbk

import xlrd
import xlwt
import pandas as pd
from xlrd import xldate_as_tuple
import datetime
import base64
# from datetime import datetime
# import datetime
'''
xlrd�е�Ԫ�����������
����һ�ɰ���������������������һ��С�������������0��1���������Ǳ����ڳ��������жϴ���ת��
��������Ҫ����������
0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error
'''


def encryption(num):
    """�����ֽ��м��ܽ��ܴ���ÿ����λ�ϵ����ֱ�Ϊ��7�˻��ĸ�λ���֣��ٰ�ÿ����λ�ϵ�����a��Ϊ10-a��"""
    if num==".":
        num=0
    newNum = []

    for i in str(num):
        if(i=="."):
            break
        if int(i):
            newNum.append(str(10 - int(i) * 7 % 10))
        else:
            newNum.append(str(0))

    # print int(''.join(newNum))
    return int(''.join(newNum))



class ExcelData():
    # ��ʼ������
    def __init__(self, data_path, sheetname):
        #����һ�����Խ����ļ�·��
        self.data_path = data_path
        # ����һ�����Խ��չ���������
        self.sheetname = sheetname
        # ʹ��xlrdģ���excel���ȡ����
        self.data = xlrd.open_workbook(self.data_path)
        # ���ݹ���������ƻ�ȡ�������е����ݣ���ʽ�٣�
        self.table = self.data.sheet_by_name(self.sheetname)
        # ���ݹ������������ȡ����������ݣ���ʽ�ڣ�
        # self.table = self.data.sheet_by_name(0)
        # ��ȡ��һ����������,���������1���ǵڶ��У������б���������
        self.keys = self.table.row_values(0)
        # ��ȡ���������Ч����
        self.rowNum = self.table.nrows
        # ��ȡ���������Ч����
        self.colNum = self.table.ncols

    # ����һ����ȡexcel��ķ���
    def readExcel(self):
        # ����һ�����б�
        datas = []
        for i in range(1, self.rowNum):
            # ����һ�����ֵ�
            sheet_data = {}
            for j in range(self.colNum):
                # print(j)
                # ��ȡ��Ԫ����������
                c_type = self.table.cell(i,j).ctype
                # print(c_type)
                # ��ȡ��Ԫ������
                c_cell = self.table.cell_value(i, j)
                if c_type==1 and j>1 and j!=4:
                    c_cell=float(c_cell)
                    # c_cell=int(c_cell)
                elif c_type == 2 and c_cell % 1 == 0:  # ���������
                    c_cell = int(c_cell)
                elif c_type == 3:
                    date = datetime.datetime(*xldate_as_tuple(c_cell, 0))
                    c_cell = date.strftime('%Y/%m/%d %H:%M:%S')
                elif c_type == 4:
                    c_cell = True if c_cell == 1 else False

                sheet_data[self.keys[j]] = c_cell
                # ѭ��ÿһ����Ч�ĵ�Ԫ�񣬽��ֶ���ֵ��Ӧ�洢���ֵ���
                # �ֵ��key����excel����ÿ�е�һ�е��ֶ�
                # sheet_data[self.keys[j]] = self.table.row_values(i)[j]
            # �ٽ��ֵ�׷�ӵ��б���
            datas.append(sheet_data)
        # ���ش�excel�л�ȡ�������ݣ����б���ֵ����ʽ����
        return datas
if __name__ == "__main__":
    data_path = "�ڶ���С��ɼ�.xls"
    sheetname = "008��"
    get_data = ExcelData(data_path, sheetname)
    order=get_data.keys
    print(order)
    datas = get_data.readExcel()
    newDatas=[]
    for i in range(2,len(datas)):

        print(datas[i])
        datas[i]["ѧ��"]=encryption(datas[i]["ѧ��"])
        datas[i]["�ܷ�"] = encryption(datas[i]["�ܷ�"])
        datas[i]["�����ܷ�"] = encryption(datas[i]["�����ܷ�"])
        datas[i]["������ܷ�"] = encryption(datas[i]["������ܷ�"])
        datas[i]["����Ƭ�α�����ܷ�"] = encryption(datas[i]["����Ƭ�α�����ܷ�"])
        datas[i]["�����1"] = encryption(datas[i]["�����1"])
        datas[i]["�����2"] = encryption(datas[i]["�����2"])
        datas[i]["����Ƭ�α����1"] = encryption(datas[i]["����Ƭ�α����1"])
        datas[i]["����Ƭ�α����2"] = encryption(datas[i]["����Ƭ�α����2"])
        datas[i]["����Ƭ�α����3"] = encryption(datas[i]["����Ƭ�α����3"])
        newDatas.append(datas[i])
    # newDatas.append(datas[0])
    # for i in range(len(datas)):
    #     flag=0
    #     for j in range(len(newDatas)):
    #         if(datas[i]['����']==newDatas[j]['����']):
    #             flag=1
    #             if(datas[i]['��������']>newDatas[j]['��������']):
    #                 temp=newDatas[j]
    #                 newDatas[j]=datas[i]
    #                 for par in order:
    #                     if(newDatas[j][par]==''and temp[par]!=''):
    #                         newDatas[j][par]=temp[par]
    #             else:
    #                 for par in order:
    #                     if (newDatas[j][par] == '' and datas[i][par] != ''):
    #                         newDatas[j][par] = datas[i][par]
    #     if(flag==0):
    #         newDatas.append(datas[i])


    pf=pd.DataFrame(list(newDatas))
    pf=pf[order]
    file_path=pd.ExcelWriter('result1.1.xls')
    pf.fillna(' ',inplace=True)
    pf.to_excel(file_path,encoding = 'utf-8',index = False)
    file_path.save()



