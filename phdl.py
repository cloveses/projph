import os
import hashlib
import xlrd
import xlsxwriter
from ph_models import *
import random,math,itertools
from exam_prog_sets import arrange_datas,PREFIX
from gen_book_tab import gen_book_tbl,count_stud_num
from gen_examid import gen_pdf

__author__ = "cloveses"

# 导入初三考生免考申请表、选考项目确认表（excel格式）到数据库中
# 二类表分别存放子目录：freeexam,itemselect之中

class MyException(Exception):
    pass

# 获取指定目录中所有文件
def get_files(directory):
    files = []
    if os.path.exists(directory):
        files = os.listdir(directory)
        files = [f for f in files if f.endswith('.xls') or f.endswith('.xlsx')]
        files = [os.path.join(directory,f) for f in files]
    return files

def compare_data(signid,phid,name):
    # 比对关键信息中考报名号、体育考号和姓名是否有误
    if count(select(s for s in StudPh if s.signid==signid and
        s.phid==phid,s.name==name)):
        return True

# 将电子表格中数据导入数据库中,同时检验数据重复和关键信息错误
@db_session
def gath_data(directory,tab_obj,convert_fun,start_row=1,
        grid_end=0,start_col=0):
    """
    start_row＝1 有一行标题行；grid_end=1 末尾1行不导入
    types       列数据类型
    start_col   从第几列开始导入
    tab_obj     导入的表模型名
    convert_fun 数据转换函数
    """
    files = get_files(directory)
    for file in files:
        print('import data from:',file)
        wb = xlrd.open_workbook(file)
        ws = wb.sheets()[0]
        nrows = ws.nrows
        for i in range(start_row,nrows-grid_end):
            datas = ws.row_values(i)
            datas = [data.strip() if isinstance(data,str) else data for 
                data in datas[start_col:]]
            datas = convert_fun(datas)
            comp_datas = []
            for k in ('signid,phid','name'):
                comp_datas.append(datas[k])
            if compare_data(comp_datas):
                if count(select(s for s in tab_obj if s.signid==datas['signid'])):
                    print('数据导入有重复，请检查：',datas,i)
                    raise MyException('数据有重复')
                else:
                    # if datas[3] in ['万浩男','孙滔','李灿灿','彭美学']:
                    #     print(datas,file)
                    tab_obj(**datas)
            else:
                print('关键信息有误：')
                print(file,'第{}行:'.format(i+1),ws.row_values(i))
                raise MyException('有关键信息错误！')


# 为所有考生设定随机值，以打乱报名号
@db_session
def set_rand():
    for s in StudPh.select(): 
        # s.sturand = random.random() * 10000 #仅2018年使用
        
        # 2019年启用以达到稳定生成准考证号（只要考试编排exam_prog_set.py
        # 不变，每次运行生成准考证号相同）
        md5_str = ''.join((s.signid,s.name,s.sex,s.idcode,s.sch,s.schcode))
        hshb = hashlib.sha3_512(md5_str.encode())
        s.sturand = hshb.hexdigest()

# 生成考生的准考证号、考试日期、所在考点
@db_session
def arrange_phid():
    # 起始考号
    phid = 1

    for arrange_data in arrange_datas:
        all_studs = []

        # 半日考试中只排某一性别，且依据随机数大小打乱报名顺序
        if len(arrange_data) == 4:
            for sch in arrange_data[-2]:
                studs = select(s for s in StudPh if s.sch==sch and
                    s.sex==arrange_data[-1]).order_by(StudPh.sturand)[:]
                all_studs.extend(studs)
        # 半日考试中同时排男女生（先女生后男生），且依据随机数大小打乱报名顺序
        elif len(arrange_data) == 3:
            for sex in ('女','男'):
                for sch in arrange_data[-1]:
                    studs = select(s for s in StudPh if s.sch==sch and
                        s.sex==sex).order_by(StudPh.sturand)[:]
                    all_studs.extend(studs)

        for stud in all_studs:
            stud.exam_addr = arrange_data[0]
            stud.exam_date = arrange_data[1]
            stud.phid = str(PREFIX + phid)
            phid +=1

def save_datas_xlsx(filename,datas):
    #将一张表的信息写入电子表格中XLSX文件格式
    w = xlsxwriter.Workbook(filename)
    w_sheet = w.add_worksheet('sheet1')
    for rowi,row in enumerate(datas):
        for coli,celld in enumerate(row):
            w_sheet.write(rowi,coli,celld)
    w.close()

# 导出各校中考报名号和体育准考证号
@db_session
def get_sch_data_xls():
    schs = select(s.sch for s in StudPh)

    tab_title = ['中考报名号','准考证号','姓名']
    for sch in schs:
        datas = [tab_title,]
        studs = select([s.signid,s.phid,s.name] for s in StudPh 
            if s.sch==sch)[:]
        datas.extend(studs)
        save_datas_xlsx(''.join((sch,'体育考号.xlsx')),datas)

#导出各学校男女考生号段
@db_session
def gen_seg_for_sch():
    datas = [['学校','女生号段','男生号段'],]
    schs = select(s.sch for s in StudPh)
    for sch in schs:
        woman_min = str(min(s.phid for s in StudPh if s.sch==sch and s.sex=='女'))
        woman_max = str(max(s.phid for s in StudPh if s.sch==sch and s.sex=='女'))
        man_min = str(min(s.phid for s in StudPh if s.sch==sch and s.sex=='男'))
        man_max = str(max(s.phid for s in StudPh if s.sch==sch and s.sex=='男'))
        datas.append([sch,'-'.join((woman_min,woman_max)),'-'.join((man_min,man_max))])
    save_datas_xlsx('各校男女考生号段.xlsx',datas)




# 检验各校上报体育选项中数据
def check_files_select(directory,types,grid_end=0,start_row=1):
    files = get_files(directory)
    if files:
        for file in files:
            infos = []
            wb = xlrd.open_workbook(file)
            ws = wb.sheets()[0]
            nrows = ws.nrows
            for i in range(start_row,nrows-grid_end):
                datas = ws.row_values(i)
                for index,(d,t) in enumerate(zip(datas,types)):
                    try:
                        if isinstance(d,str):
                            d = d.strip()
                        if d != '':
                            t(d)
                    except:
                        # print(datas)
                        infos.append('文件：{}中，第{}行，第{}列数据有误'.format(file,i+1,index+1))
            if infos:
                print(file,'数据格式有误！')
                print(infos)
        print('检验的目录：',directory)
        # if infos:
        #     for info in infos:
        #         print(info)
        #选项校验
        if not infos:
            for file in files:
                wb = xlrd.open_workbook(file)
                ws = wb.sheets()[0]
                nrows = ws.nrows
                for i in range(start_row,nrows-grid_end):
                    datas = ws.row_values(i)
                    datas = [i.replace(' ','') if isinstance(i,str) else i for i in datas[-4:]]
                    selects = [int(i) if i else 0 for i in datas]
                    if not (sum(selects) == 0 or (selects[0]+selects[1] == 1 and selects[-2]+selects[-1]==1)):
                        infos.append('文件：{}中，第{}行选项有误'.format(file,i+1))
        if infos:
            for info in infos:
                print(info)
        else:
            print('检验通过！')

# 检验各校上报的体育免考生数据
def check_files_other(directory,types,grid_end=0,start_row=1):
    print('检验的目录：',directory)
    files = get_files(directory)
    if files:
        for file in files:
            infos = []
            wb = xlrd.open_workbook(file)
            ws = wb.sheets()[0]
            nrows = ws.nrows
            for i in range(start_row,nrows-grid_end):
                datas = ws.row_values(i)
                for index,(d,t) in enumerate(zip(datas,types)):
                    try:
                        if isinstance(d,str):
                            d = d.strip()
                        if d != '':
                            t(d)
                    except:
                        infos.append('文件：{}中，第{}行，第{}列数据有误'.format(file,i+1,index+1))
            if infos:
                # for info in infos:
                #     print(info)
                print(file,'错误！')
            else:
                print(file,'数据检验通过！')

# 获取指定中考报名号学生的所在学校
@db_session
def get_sch(signid):
    stud = select(s for s in StudPh if s.signid == signid).first()
    if stud:
        return stud.sch
    else:
        '没有查到该生所在学校。'

# 检验选择错误
@db_session
def check_select():
    for stud in ItemSelect.select():
        if (stud.jump_option + stud.rope_option + stud.globe_option +
                    stud.bend_option) == 0:
            if count(FreeExam.select(lambda s:s.signid == stud.signid)) != 1:
                print(stud.signid,stud.name,get_sch(stud.signid),'未免考考生无选项！')
        else:
            if not (stud.jump_option + stud.rope_option == 1 and 
                    stud.globe_option + stud.bend_option == 1):
                print(stud.signid,stud.name,get_sch(stud.signid),'选项有误，请检查！')

# 导入考生选项表至总表StudPh
@db_session
def put2studph():
    for stud in ItemSelect.select():
        studph = select(s for s in StudPh if s.signid == stud.signid).first()
        if not studph:
            print(stud.signid,stud.name,get_sch(stud.signid),'考号错误，查不到此人！')
        else:
            if stud.jump_option + stud.rope_option + stud.globe_option + stud.bend_option == 0:
                studph.free_flag = True
            else:
                studph.free_flag = False
            studph.set(jump_option=stud.jump_option,
                rope_option=stud.rope_option,
                globe_option=stud.globe_option,
                bend_option=stud.bend_option)

# 导出各校体育考试确认表
@db_session
def dump_itemselect_for_sch():
    schs = select(s.sch for s in StudPh)
    for sch in schs:
        all_signid = select(s.signid for s in StudPh if s.sch==sch)
        studs = select((s.signid,s.phid,s.name,s.jump_option,
            s.rope_option,s.globe_option,s.bend_option) for s in ItemSelect if s.signid in all_signid)[:]
        datas = [['序号','中考报名号','准考证号','姓名','立定跳远','跳绳','实心球','体前屈'],]
        datas.extend(studs)
        save_datas_xlsx(sch+'确认表.xlsx',datas)

if __name__ == '__main__':
    # print('注意：执行时应将有关字体文件放入当前目录中')
    # print('''执行前所有数据导入与生成要具备两个条件：
    #     1.有要导入的考生信息表(在 studph 子目录中,注意字段顺序)，
    #     2.exam_prog_sets.py 文件中有考试日程安排信息和准考证前缀码
    #     ''')
    # exe_flag = input('是否执行前期所有数据导入与生成(y/n)：')
    # if exe_flag == 'y':

    #     exe_flag = input('是否执行考生信息导入(y/n)：')
    #     if exe_flag == 'y':
    #         ensure = input('ensure:')
    #         if ensure == 'y':
    #             gath_data(StudPh,STUDPH_KS,'studph',0,types=STUDPH_TYPE)

    #     exe_flag = input('是否执行添加用于生成考号的随机数(y/n)：')
    #     if exe_flag == 'y':
    #         ensure = input('ensure:')
    #         if ensure == 'y':
    #             set_rand()

    #     exe_flag = input('是否执行生成考生准考证号并安排考点和考试时间(y/n)：')
    #     if exe_flag == 'y':
    #         ensure = input('ensure:')
    #         if ensure == 'y':
    #             arrange_phid()

    #     exe_flag = input('是否执行导出各校考生报名号和准考证号(y/n)：')
    #     if exe_flag == 'y':
    #         get_sch_data_xls()

    #     exe_flag = input('是否执行生成各校男女考生准考证号段(y/n)：')
    #     if exe_flag == 'y':
    #         gen_seg_for_sch() #排理化实验用

    #     exe_flag = input('是否执行生成各考点异常登记表(y/n)：')
    #     if exe_flag == 'y':
    #         datas = gen_book_tbl()
    #         save_datas_xlsx('各时间段各考点考试分组号.xlsx',datas)

    #     exe_flag = input('是否执行生成各时间段各考点考生人数(y/n)：')
    #     if exe_flag == 'y':
    #         datas = count_stud_num()
    #         save_datas_xlsx('各时间段各考点考生人数.xlsx',datas)

    # # print('''
    # #     生成准考证,照片文件子目录为pho：
    # #     ''')
    # # exe_flag = input('启动生成准考证？(y/n)：')
    # # if exe_flag == 'y':
    # #     gen_all_examid()


    # print('''
    #     要检验的免试表和选项表应分别存放于以下子目录中：
    #     freeexam
    #     itemselect
    #     ''')
    # exe_flag = input('启动检验免试表xls数据和选项表xls数据(y/n)：')
    # if exe_flag == 'y':
    #     check_files_other('freeexam',FREE_EXAM_TYPE)
    #     check_files_select('itemselect',ITEM_SELECT_TYPE)
        
    # exe_flag = input('免试表xls和选项表xls导入到数据库中，验证后放在studph中(y/n)：')
    # if exe_flag == 'y':
    #     gath_data('freeexam',FreeExam,convert_freeexam_data)
    #     gath_data('itemselect',ItemSelect,convert_itemselect_data)
    #     check_select()
    #     put2studph()


    # check_files_other('freeexam',FREE_EXAM_TYPE)
    # check_files_select('itemselect',ITEM_SELECT_TYPE)
    # # 导入体育选项表
    # gath_data_itemselect(ItemSelect,ITEM_SELECT_KS,'itemselect',0,types=ITEM_SELECT_TYPE,check_repeat=True) # 末尾行无多余数据
    # # 检查所有免表
    # check_files_other('freeexam',FREE_EXAM_TYPE)
    # # 导入免试表
    # gath_data(FreeExam,FREE_EXAM_KS,'freeexam',0,types=FREE_EXAM_TYPE)
    # # 分校导出确认表
    # # dump_itemselect_for_sch()
    # put2studph()
    # check_select()