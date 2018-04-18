import os
import math
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.units import mm
from reportlab.graphics.barcode import code39, code128, code93
from ph_models import *

# 2018使用 每页A4纸上打印8张准考证

## 尺寸mm
ID_SIZE = (210,297)
## 水印文本
WATERMARK_TXT = "泗县教体局"
# 标题及其位置
TITLE = (5*mm,74*mm,"2018年初中学业水平考试")
ID_NAME = (10*mm,64*mm,"准　考　证")
## mm
POS_X = 5 #(5,105+5)
POS_Y = (54,44,34,24,14)
## 以上为以下七项的输出位置
ITEM_NAMES = ("准考号:  ","姓　名:  ","性　别:  ","考　点:  ","报名点:  ")

STUDS = [('18251402',"李文娟",'男','泗县一中','01中学','5000','183222508209'),
]
IMG_PATH = ".\\idsd"

BAR_METHODS = {'code39':code39.Extended39, 
            'code128':code128.Code128,
            'code93':code93.Standard93}

# 条形码打印位置
BAR_X = 50
BAR_Y = 14

# 照片打印位置
IMG_X = 50
IMG_Y = 22

pos_trans = ((0,0),(0,72),(0,72*2),(0,72*3),
    (105,0),(105,72),(105,72*2),(105,72*3))

def confirm_path(path):
    if not os.path.exists(path):
        os.makedirs(path)

def set_font(canv,size,font_name='msyh',font_file='msyh.ttf'):
    pdfmetrics.registerFont(TTFont(font_name,font_file))
    canv.setFont(font_name,size)

def draw_barcode(canv,idcode,codetype='code128'):
    ## 绘制条形码
    ## codetype have:code39,code93,code128
    barcd = BAR_METHODS[codetype](idcode,barWidth=1,humanReadable=True)
    barcd.drawOn(canv,BAR_X*mm,BAR_Y*mm)
    set_font(canv,16,font_name='simsun',font_file='simsun.ttc')
    canv.drawString(BAR_X*mm+10,BAR_Y*mm,"‡")
    canv.drawString(BAR_X*mm+34.2*mm,BAR_Y*mm,"‡")
    # barcode39 = code39.Extended39('34322545666',barHeight=1*cm,barWidth=0.8)
    # barcode39.drawOn(c,20,20)
    # barcode93 = code93.Standard93('34322545666')
    # barcode93.drawOn(c,20,60)
    # barcode128 = code128.Code128('34322545666')
    # barcode128.drawOn(c,20,100)

def draw_one(canv,pos_trans,stud):
    canv.saveState()
    canv.translate(pos_trans[0]*mm,pos_trans[1]*mm)
    # 绘制标题
    set_font(canv,11)
    canv.drawString(*TITLE)
    set_font(canv,18)
    canv.drawString(*ID_NAME)
    set_font(canv,10)
    ## 背景图
    # canv.drawImage('bg.jpg',POSITIONS[-1][0]*mm,POSITIONS[-1][1]*mm)
    for index,info in enumerate(stud[:5]):
        canv.drawString(POS_X*mm,POS_Y[index]*mm,''.join((ITEM_NAMES[index],info)))
    # 绘制照片
    file_name = ''.join(('Z',stud[-1],'.jpg'))
    path = os.path.join('photos',stud[-2],file_name)
    canv.drawImage(path,IMG_X*mm,IMG_Y*mm,width=30*mm,height=30*1.32*mm)
    # 绘制条形码
    draw_barcode(canv,stud[0])
    # 绘制水印
    set_font(canv,10)
    canv.setFillColorRGB(180,180,180,alpha=0.3)
    canv.drawString(IMG_X*mm+mm,IMG_Y*mm+0.5*mm,WATERMARK_TXT)
    canv.restoreState()
    # canv.showPage()

def draw_page(canv,studs):
    for pos,stud in zip(pos_trans,studs):
        draw_one(canv,pos,stud)
    canv.showPage()

def gen_pdf(dir_name,sch_name,studs,page_num=8):
    confirm_path(dir_name)
    path = os.path.join(dir_name,sch_name + '.pdf')
    canv = canvas.Canvas(path,pagesize=(ID_SIZE[0]*mm,ID_SIZE[1]*mm))
    pages = math.ceil(len(studs)/page_num)
    for i in range(pages):
        draw_page(canv,studs[i*page_num:(i+1)*page_num])
    canv.save()

# def gen_all_pdfs(dir_name):
#     schs = getdata.get_schs()
#     for sch in schs:
#         studs = getdata.get_studs(sch)
#         gen_pdf(dir_name,sch,studs)


#以下各函数中照片文件未生成

# 按学校生成准考证
@db_session
def gen_examid_sch(dir_name):
    schs = select(s.sch for s in StudPh)
    for sch in schs:
        datas = select(s 
         for s in StudPh if s.sch==sch).order_by(StudPh.classcode,StudPh.phid)
        datas = [(s.phid,s.name,s.sex,s.exam_addr,s.sch,s.schcode,s.signid) for s in datas]
        gen_pdf(dir_name,sch,datas)

# # 按时间段（半日）和考点生成准考证
@db_session
def gen_bak_examid(dir_name):
    exam_addrs = select(s.exam_addr for s in StudPh)
    exam_dates = select(s.exam_date for s in StudPh)
    for exam_date in exam_dates:
        for exam_addr in exam_addrs:
            studs = select(s for s in StudPh if s.exam_addr==exam_addr and 
                s.exam_date==exam_date).order_by(StudPh.phid)
            studs = [(s.phid,s.name,s.sex,s.exam_addr,s.sch,s.schcode,s.signid) for s in studs]
            gen_pdf(dir_name,exam_addr+exam_date,studs)

if __name__ == '__main__':
    # gen_pdf('idsd','泗县草沟中学',STUDS*8)
    # gen_examid_sch('idsd')
    gen_bak_examid('bakid')