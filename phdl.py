import os
import hashlib
import re
import xlrd
import xlsxwriter
from ph_models import *
import random,math,itertools
from exam_prog_sets import arrange_datas,PREFIX,date_seq,sch_seq
from gen_book_tab import gen_book_tbl,count_stud_num,PAGE_NUM

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

@db_session
def compare_data(signid,phid,name):
    # 比对关键信息中考报名号、体育考号和姓名是否有误
    if count(select(s for s in StudPh if s.signid==signid and
        s.phid==phid and s.name==name)):
        return True

@db_session
def init_studph_data(directory='studph',start_row=1):
    files = get_files(directory)
    for file in files:
        print('import data from:',file)
        wb = xlrd.open_workbook(file)
        ws = wb.sheets()[0]
        nrows = ws.nrows
        for i in range(start_row,nrows):
            datas = ws.row_values(i)
            datas = [str(int(data)) if isinstance(data,float) else data
                for data in datas]
            params = {}
            for k,v in zip(STUDPH_KS,datas):
                params[k] = v
            StudPh(**params)

# 将电子表格中数据导入数据库中,同时检验数据重复和关键信息错误
@db_session
def gath_data(directory,start_row=1,grid_end=0,start_col=0):
    """
    start_row＝1 有一行标题行；grid_end=1 末尾1行不导入
    types       列数据类型
    start_col   从第几列开始导入
    tab_obj     导入的表模型名
    convert_fun 数据转换函数
    """
    if directory in ('freeexam','addfreeexam'):
        tab_obj = FreeExam
        convert_fun = convert_freeexam_data
    elif directory == 'itemselect':
        tab_obj = ItemSelect
        convert_fun = convert_itemselect_data
    else:
        raise MyException('目录名错误！')
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
            for k in ('signid','phid','name'):
                comp_datas.append(datas[k])
            if compare_data(*comp_datas):
                if count(select(s for s in tab_obj if s.signid==datas['signid'])):
                    print('数据导入有重复，请检查：',datas,i)
                    raise MyException('数据有重复')
                else:
                    # if datas[3] in ['万浩男','孙滔','李灿灿','彭美学']:
                    #     print(datas,file)
                    tab_obj(**datas)
                    # 如果为后补免试表的，同时修改选项表和学生成绩总表
                    if directory == 'addfreeexam':
                        data_sets = dict(
                            jump_option = 0,
                            rope_option = 0,
                            globe_option = 0,
                            bend_option = 0)
                        # 更新选项表
                        stud = select(s for s in ItemSelect if s.signid==datas['signid']).first()
                        stud.set(**data_sets)
                        # 更新体育成绩总表
                        data_sets['free_flag'] = True
                        stud = select(s for s in StudPh if s.signid==datas['signid']).first()
                        stud.set(**data_sets)
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

        # 半日考试中只排某一性别，且依据班级分组（同一班级尽量同组考试）
        if len(arrange_data) == 4:
            for sch in arrange_data[-2]:
                studs = select(s for s in StudPh if s.sch==sch and
                    s.sex==arrange_data[-1]).order_by(StudPh.classcode)[:]
                all_studs.extend(studs)
        # 半日考试中同时排男女生（先女生后男生），且依据班级分组（同一班级尽量同组考试）
        elif len(arrange_data) == 3:
            for sex in ('女','男'):
                for sch in arrange_data[-1]:
                    studs = select(s for s in StudPh if s.sch==sch and
                        s.sex==sex).order_by(StudPh.classcode)[:]
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

# 导出各考点分组－－号段对照表，便于检录查找学生所属组数
# 随异常情况登记表发放各考点检录处
@db_session
def group_phid_arrange():
    sexes = ('女','男')
    group_seq = 1
    for exam_addr in sch_seq:
        datas = [['考试日期','性别','组数','起始号段','结束号段'],]
        for dateseq in date_seq:
            for sex in sexes:
                studs = select(s for s in StudPh if s.exam_addr==exam_addr and
                    s.sex==sex and s.exam_date==dateseq).order_by(StudPh.phid)[:]
                totals = len(studs)
                unit_group_num = math.ceil(totals/PAGE_NUM)
                for i in range(unit_group_num):
                    temp_studs = studs[i * PAGE_NUM:(i+1)*PAGE_NUM]
                    datas.append([dateseq,sex,group_seq,temp_studs[0].phid,temp_studs[-1].phid])
                    group_seq += 1
        save_datas_xlsx(''.join((exam_addr,'分组-号段对照表','.xlsx')),datas)


# 检验体育选项是否符合要求
def check_item_select(file,datas,row_number):
    datas = [i.replace(' ','') if isinstance(i,str) else i for i in datas[-4:]]
    selects = [int(i) if i else 0 for i in datas]
    if not (sum(selects) == 0 or (selects[0]+selects[1] == 1 and selects[-2]+selects[-1]==1)):
        return '文件：{}中，第{}行选项有误'.format(file,row_number)

#检验体育选项表一行的数据类型
def check_item_select_types(file,datas,row):
    for index,(d,t) in enumerate(zip(datas,ITEM_SELECT_TYPE)):
        try:
            if isinstance(d,str):
                d = d.strip()
            if d != '':
                t(d)
        except:
            return '文件：{}中，第{}行，第{}列数据有误'.format(file,row,index+1)

# 检验免试表一行的数据类型
def check_free_exam_types(file,datas,row):
    for index,(d,t) in enumerate(zip(datas,FREE_EXAM_TYPE)):
        try:
            if isinstance(d,str):
                d = d.strip()
            if d != '':
                t(d)
        except:
            return '文件：{}中，第{}行，第{}列数据有误'.format(file,row,index+1)

# 检验xls文件中数据
# 检验各校上报体育选项和免试名单xls文件中数据
# 只要指定固定的目录名即可：freeexam,itemselect

def check_files(directory,start_row=1,grid_end=0):
    if directory  in ('freeexam','addfreeexam'):
        check_types = check_free_exam_types
        col_num = 7
        check_fun = None
    elif directory == 'itemselect':
        check_types = check_item_select_types
        col_num = 8
        check_fun = check_item_select
    else:
        raise MyException('目录名错误！')
    files = get_files(directory)
    print('\n','检验的目录：',directory)
    print('............................')
    if files:
        for file in files:
            infos = []
            wb = xlrd.open_workbook(file)
            ws = wb.sheets()[0]
            nrows = ws.nrows
            if ws.ncols != col_num:
                print('数据列数不符合要求:',file,'\n')
                continue
            for i in range(start_row,nrows-grid_end):
                datas = ws.row_values(i)
                info = check_types(file,datas,i+1)
                if info:
                    infos.append('文件：{}中，第{}行，第{}列数据有误'.format(file,i+1,index+1))
                if len(infos) >= 3:
                    print('数据类型错误：',file,'\n')
                    for info in infos:
                        print(info)
                    break
            # 校验行数据
            if not infos and check_fun is not None:
                for i in range(start_row,nrows-grid_end):
                    info = check_fun(file,ws.row_values(i),i+1)
                    if info:
                        infos.append(info)

            # 检验关键信息
            for i in range(start_row,nrows-grid_end):
                keyinfos = ws.row_values(i)[1:4]
                keyinfos = (str(int(keyinfos[0])),str(int(keyinfos[1])),keyinfos[2])
                if not compare_data(*keyinfos):
                    print('关键信息检验失败：')
                    print(file,'第{}行'.format(i+1))
                    print(ws.row_values(i))
                    print('*******************************')
            if infos:
                print()
                print('检验失败：',file,'\n')
                for info in infos:
                    print(info)
            else:
                print('检验通过：',file,'\n')
    print('------------------------')


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
                print(stud.signid,stud.name,get_sch(stud.signid),'数据错误：未免考考生无体育选项！')
        elif not (stud.jump_option + stud.rope_option == 1 and 
                    stud.globe_option + stud.bend_option == 1):
                print(stud.signid,stud.name,get_sch(stud.signid),'选项有误，请检查！')
        else:
            if count(FreeExam.select(lambda s:s.signid == stud.signid)) >= 1:
                print(stud.signid,stud.name,get_sch(stud.signid),'数据错误：该生有体育选项，也有免试！')

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
        datas = [['中考报名号','准考证号','姓名','立定跳远','跳绳','实心球','体前屈'],]
        datas.extend(studs)
        save_datas_xlsx(sch+'确认表.xlsx',datas)

# 导出免试总表
@db_session
def dump_freeexam_studs():
    studs = select(s for s in FreeExam)
    studs = [(s.signid,s.phid,s.name,s.reason,s.material,s.memo) for s in studs]
    datas = [['中考报名号','准考证号','姓名','免考原因','申请材料','备注'],]
    datas.extend(studs)
    save_datas_xlsx('全县免考表.xlsx',datas)

# 导入免试类型为1的考生 1 满分
@db_session
def freexam_type2studph(file='全县免考满分表.xlsx'):
    wb = xlrd.open_workbook(file)
    ws = wb.sheets()[0]
    for i in range(ws.nrows):
        datas = ws.row_values(i)
        if isinstance(datas[-1],float):
            freetype = int(datas[-1])
            stud = select(s for s in StudPh if s.signid==str(int(datas[0]))).first()
            if stud:
                stud.freetype = freetype
            else:
                print('无该考生：',int(datas[0]),int(datas[1]),datas[2])
        else:
            print(int(datas[0]),int(datas[1]),datas[2],'数据类型错误')

# 将跑步时间字符串转换为   秒数*10
def chg_run_data(data):
    m,s = data.split(':')
    ms = int(m) * 60
    return int((ms + float(s)) * 10)

# 首先手工处理EXCEL表中的错误，之后用以下方法来进行检查
# 检查EXCEL表中的成绩数据 总分累计是否正确 赋分是否有问题 是否有缺项
@db_session
def check_scores(file='2018体育考试成绩汇总表.xls'):
    wb = xlrd.open_workbook(file)
    ws = wb.sheets()[0]
    infos = []
    # 检验考试数据错误，缺少项目成绩及分数,四项无成绩的,分数计算有误的
    for i in range(1,ws.nrows):
        datas = ws.row_values(i)
        # print(i,datas)
        phid = str(int(datas[0]))
        name = datas[2]
        sex = datas[3]
        sch = datas[5]
        info = []
        total_score = int(datas[4])
        totals = 0
        stud = select(s for s in StudPh if s.phid==phid).first()
        for j in range(6,18):
            if isinstance(datas[j],str):
                datas[j] = datas[j].strip()
        if not stud.free_flag:
            if total_score == 0:
                info.append('全无成绩，可能为缺考！')
            else:
                if not (datas[6] or datas[16]):
                    info.append('跳远和跳绳同时没有成绩！')

                if datas[6] and datas[16]:
                    info.append('跳远和跳绳不能同时有成绩！')
                else:
                    if datas[6] != '':
                        data = int(datas[6])
                        score = int(datas[7])
                        totals += score
                        if sex == '女':
                            look_score = score_jump_woman(data)
                        else:
                            look_score = score_jump_man(data)
                        if score != look_score:
                            info.append('跳远赋分错误')

                    if datas[16] != '':
                        data = int(datas[16])
                        score = int(datas[17])
                        totals += score
                        if sex == '女':
                            look_score = score_rope_woman(data)
                        else:
                            look_score = score_rope_man(data)
                        if score != look_score:
                            info.append('跳绳赋分错误')

                if not (datas[8] or datas[14]):
                    info.append('体前屈和实心球同时没有成绩！')

                if datas[8] and datas[14]:
                    info.append('体前屈和实心球不能同时有成绩！')
                else:
                    if datas[8] != '':
                        data = int(float(datas[8])*10)
                        score = int(datas[9])
                        totals += score
                        if sex == '女':
                            look_score = score_bend_woman(data)
                        else:
                            look_score = score_bend_man(data)
                        if score != look_score:
                            info.append('体前屈赋分错误')

                    if datas[14] != '':
                        data = int(float(datas[14])*100)
                        score = int(datas[15])
                        totals += score
                        if sex == '女':
                            look_score = score_globe_woman(data)
                        else:
                            look_score = score_globe_man(data)
                        if score != look_score:
                            info.append('实心球赋分错误')

                if not (datas[10] or datas[12]):
                    info.append('该生无跑步成绩！')

                #计算跑步成绩
                if isinstance(datas[11],float) or datas[11]:
                    totals += int(datas[11])
                if isinstance(datas[13],float) or datas[13]:
                    totals += int(datas[13])

                if datas[10] and datas[12]:
                    info.append('800跑和1000跑不能同时有成绩！')
                else:
                    if datas[10] != '':
                        if isinstance(datas[10],str):
                            if ':' not in datas[10]:
                                print(i+1,phid,'10数据格式错误！')
                            else:
                                data = int(chg_run_data(datas[10]))
                                score = int(datas[11])
                                look_score = score_run_woman(data)
                                if score != look_score:
                                    info.append('800跑赋分错误{} {}'.format(datas[10],score))
                        else:
                            print(i+1,phid,'10数据格式错误！')

                    if datas[12] != '':
                        if isinstance(datas[12],str):
                            if ':' not in datas[12]:
                                print(i+1,phid,'12数据格式错误！')
                            else:
                                data = int(chg_run_data(datas[12]))
                                score = int(datas[13])
                                look_score = score_run_man(data)
                                if score != look_score:
                                    info.append('1000跑赋分错误{} {}'.format(datas[12],score))
                        else:
                            print(i+1,phid,'12数据格式错误！')

                if total_score != totals:
                    info.append('总分有误！')
            if info:
                addinfo = [phid,name,sch]
                addinfo.extend(info)
                infos.append(addinfo)
    if infos:
        for info in infos:
            print(info)

def clear_test_data(datas):
    datas[0] = str(int(datas[0])) if isinstance(datas[0],float) or datas[0] else ''
    datas[2] = str(int(float(datas[2]) * 10)) if isinstance(datas[2],float) or datas[2] else ''
    datas[8] = str(int(float(datas[8]) * 100)) if isinstance(datas[8],float) or datas[8] else ''
    datas[10] = str(int(datas[10])) if isinstance(datas[10],float) or datas[10] else ''
    for i in range(1,12,2):
        datas[i] = int(datas[i]) if datas[i] != '' else ''
    return datas

# 导入经过手工处理的正确的成绩汇总表
@db_session
def score2studph(file='2018体育考试成绩汇总表.xls'):
    wb = xlrd.open_workbook(file)
    ws = wb.sheets()[0]

    for i in range(1,ws.nrows):
        datas = ws.row_values(i)
        params = {}
        phid = str(int(datas[0]))
        cardid = str(int(datas[1]))
        name = datas[2]
        sex = datas[3]
        total_score = int(datas[4])
        params['cardid'] = cardid
        params['total_score'] = total_score
        # 清理字符串中的空格
        for j in range(6,18):
            if isinstance(datas[j],str):
                datas[j] = datas[j].strip()

        datas_right = clear_test_data(datas[6:])
        for k,v in zip(SCORE_KS,datas_right):
            if v != '':
                params[k] = v

        stud = select(s for s in StudPh if s.phid==phid).first()
        if stud and stud.name==name:
            stud.set(**params)

def chg_key(s,k):
    # 导出成绩时，测试数据中有例外情况（如弃考等成绩为0的），则导出例外情况
    if k.endswith('score') and not k.startswith('total'):
        data_k = '_'.join((k.split('_')[0],'data'))
        v = getattr(s,data_k)
        if v:
            if not re.match(r'\d+',v):
                return data_k
    return k

# 依据要保存的学生和字段名获取信息
def get_score_data(studs,keys):
    datas=[]
    # datas.append(keys)
    for s in studs:
        data = [getattr(s,chg_key(s,k)) for k in keys]
        datas.append(data)
    return datas
#导出所有学生的体育考试成绩 总体导出和分校导出
@db_session
def dump_score():
    col_name_all = ('signid','phid','name','sch','schcode','total_score')
    col_name_sch = ('signid','phid','name','sch','schcode','classcode','total_score',
        'jump_score','rope_score','globe_score','bend_score','run8_score','run10_score')
    studs = select(s for s in StudPh).order_by(StudPh.phid)
    datas = [['中考报名号','体育考试号','姓名','学校','学校代码','总分'],]
    datas.extend(get_score_data(studs,col_name_all))
    save_datas_xlsx('全县体育考试分数.xlsx',datas)

    schs = select(s.sch for s in StudPh)
    for sch in schs:
        studs = select(s for s in StudPh if s.sch==sch).order_by(StudPh.classcode)
        datas = [['中考报名号','体育考试号','姓名','学校','学校代码','班级代码','总分',
            '立定跳远','跳绳','实心球','体前屈','800米跑','1000米跑'],]
        datas.extend(get_score_data(studs,col_name_sch))
        save_datas_xlsx(sch+'体育考试成绩.xlsx',datas)

TOTAL_SCORE = 60

@db_session
def set_freeexam_score():
    for stud in select(s for s in StudPh if s.free_flag==True):
        if stud.freetype == 1:
            stud.total_score = TOTAL_SCORE
        else:
            stud.total_score = int(TOTAL_SCORE * 0.6)

# 初始化数据库表
@db_session
def init_tab(tab_objs):
    '''tab_objs 模型列表'''
    for tab_obj in tab_objs:
        delete(s for s in tab_obj)

#后补免试表检查并导入数据库表
@db_session
def add_freeexam():
    print('...检查免试表文件...')
    check_files('addfreeexam')
    if input('无错直接回车则导入至数据库中选项表和总表：') == '':
        print('...导入免试表...')
        gath_data('addfreeexam')
        print('...开始检查数据...')
        check_select()


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
    #             init_studph_data()

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
        # 导出分组－考号对照表
        # group_phid_arrange()
    # print('''
    #     要检验的各校上报的免试表存放子目录中：
    #     freeexam
    #     选项表应分别存放于子目录中:
    #     itemselect
    #     导入数据后再补的免试表存放子目录中：
    #     addfreeexam
    #     ''')
    # exe_flag = input('免试表xls文件和选项表xls文件检验(y/n)：')
    # if exe_flag == 'y':
    #     check_files('freeexam')
    #     check_files('itemselect')
        
    # exe_flag = input('全部免试表xls和选项表xls导入到数据库中并检验(y/n)：')
    # if exe_flag == 'y':
    #     print('初始化数据库表...')
    #     init_tab([FreeExam,ItemSelect])
    #     print('...初始化完成！')
    #     print('...导入免试表...')
    #     gath_data('freeexam')
    #     print('...导入选项表...')
    #     gath_data('itemselect')
    #     print('...开始检查数据...')
    #     check_select()

    # exe_flag = input('是否启动数据合并到正式表StudPh中：(y/n)')
    # if exe_flag == 'y':
    #     # 数据合并到正式表StudPh中
    #     put2studph()

    # exe_flag = input('是否启动添加和导入后补免试考生：(y/n)')
    # if exe_flag == 'y':
    #     add_freeexam() # 添后补免试考生 （检查文件、改入数据库和修改选项）

    # # dump_itemselect_for_sch()
    # # dump_freeexam_studs() #导出全县免试表

    # exe_flag = input('是否启动免试考生赋分 满分表导入（满分和60%分）：(y/n)')
    # if exe_flag == 'y':
    #     freexam_type2studph() #从文件 全县免考满分表.xlsx导入免试类型至总表 
    #     set_freeexam_score()  #免考学生赋分
    #对成绩的电子表格文件进行检查
    # check_scores('score\\2018泗县体考成绩上报.xls')
    #将正确的成绩从电子表格文件中导入数据库
    score2studph('score\\2018泗县体考成绩上报.xls')
    #导出学生的体育成绩（全县版和分校版）
    # dump_score()