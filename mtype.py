import xlrd

__author__ = 'cloveses'

# 本程序用于读取电子表格数据时对指定的单元格转换为指定的数据类型
# 用于导入电子表格中的数据至数据库中

TYPES = ('empty','string','number','date','boolean','error')

def TypeExcetion(Exception):
    pass

# 用偏函数的形式来固定非本程序所用的默认返回值
# functools.partial(mstr,default=None)
# eg:
# mtype : cell_type(i,j)

def mstr(data,mtype,default=None):
    # blank 0,return default
    # str   1,return immediately
    if mtype == 0:
        return default

    if mtype == 1:
        return data.strip()
    # float 2,cut .0
    if mtype == 2:
        data = str(data)
        if data.endswith('.0'):
            data = data[:-2]
        return data

    raise TypeExcetion('Type Exception:\nExpect string but {}(type is:{})'
        .format(data,TYPES[mtype]))

def mfloat(data,mtype,default=None):
    if mtype == 0:
        return default
    if mtype == 2:
        return data
    if mtype == 1:
        try:
            data = float(data)
        except:
            raise TypeExcetion('Type Exception:\nExpect float but {}'
                .format(data))
        else:
            return data
    raise TypeExcetion('Type Exception:\nExpect float but {}(type is:{})'
        .format(data,TYPES[mtype]))

def mint(data,mtype,default=None):
    if mtype == 0:
        return default
    if mtype in (1,2):
        try:
            data = int(data)
        except:
            raise TypeExcetion('Type Exception:\nExpect int but {}'
                .format(data))
        else:
            return data

    raise TypeExcetion('Type Exception:\nExpect int but {}(type is:{})'
        .format(data,TYPES[mtype]))

def mdate(data,mtype,default=None,datemode=0):
    # datemode 0 : based 1900 (available:ws.datemode)
    # datemode 1 : based 1904
    if mtype == 3:
        data = xlrd.xldate.xldate_as_datetime(wb.cell_value(9,0),datemode)
        return data
    if mtype == 0:
        return default
    raise TypeExcetion('Type Exception:\nExpect date but {}(type is:{})'
        .format(data,TYPES[mtype]))

def mbool(data,mtype,default=None):
    if mtype == 4:
        if data:
            return True
        else:
            return False
    # blank space
    if mtype == 0:
        return default

    raise TypeExcetion('Type Exception:\nExpect boolean but {}(type is:{})'
        .format(data,TYPES[mtype]))
