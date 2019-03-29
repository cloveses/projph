import os
import datetime

import xlrd, xlsxwriter
import string



def get_files(directory):
    files = []
    files = os.listdir(directory)
    files = [f for f in files if f.endswith('.xls') or f.endswith('.xlsx')]
    files = [os.path.join(directory,f) for f in files]
    return files


def get_datas(filename, start=1, date_col_num=None, int_col_num=None, float_col_num=None, date_deal_fun=None):
    datas = []
    if not date_col_num:
        date_col_num = []
    if not int_col_num:
        int_col_num = []
    if not float_col_num:
        float_col_num = []

    w = xlrd.open_workbook(filename)
    datemode = w.datemode
    ws = w.sheets()[0]
    datas.append(ws.row_values(0))
    nrows = ws.nrows
    for rowi in range(start, nrows):
        primitive_datas = []
        for coli,celld in enumerate(ws.row_values(rowi)):
            ctype = ws.cell(rowi, coli).ctype
            print(rowi, coli, ctype, celld)
            if ctype == 0:
                primitive_datas.append('')
                continue
            if ctype == 1:
                celld = celld.strip()
            if coli in int_col_num:
                try:
                    d = int(celld)
                except Exception as e:
                    print(e)
                    d = ''
            elif coli in float_col_num:
                try:
                    d = float(celld)
                except Exception as e:
                    print(e)
                    d = ''
            elif coli in date_col_num:
                if ctype == 3:
                    d = xlrd.xldate.xldate_as_datetime(celld, datemode)
                else:
                    if date_deal_fun:
                        d = date_deal_fun(celld)
                    else:
                        d = ''
            else:
                if ctype == 2:
                    dstr = str(celld)
                    if int(dstr[dstr.index('.')+1:]) == 0:
                        d = str(int(celld))
                    else:
                        d = dstr
                else:
                    d = str(celld)
            primitive_datas.append(d)
        datas.append(primitive_datas)
    return datas

def save_datas_xlsx(filename,datas):
    #将一张表的信息写入电子表格中XLSX文件格式
    w = xlsxwriter.Workbook(filename)
    date_format = w.add_format({'num_format':'yyyy/mm/dd'})
    w_sheet = w.add_worksheet('sheet1')
    for rowi,row in enumerate(datas):
        for coli,celld in enumerate(row):
            if isinstance(celld, datetime.datetime):
                w_sheet.write(rowi,coli,celld, date_format)
            else:
                w_sheet.write(rowi,coli,celld)
    w.close()

def check_idcode(stud):
    '''身份证号验证程序'''
    idcode = stud.idcode
    if idcode:
        checkcodes = ['1', '0', 'X','9', '8', '7', '6', '5', '4', '3', '2']
        wi = [7,9,10,5,8,4,2,1,6,3,7,9,10,5,8,4,2]
        s = sum((i*int(j) for i,j in zip(wi,idcode[:-1])))
        checkcode = checkcodes[s % 11]
        # ck = idcode[-1].upper()
        ck = idcode[-1]
        if checkcode != ck:
            return '校验出错！'
        if int(idcode[-2]) % 2 == 0 and stud.sex == '男':
            return '性别码出错！'
        if int(idcode[-2]) % 2 == 1 and stud.sex == '女':
            return '性别码出错！'
    else:
        return '身份证号为空'
