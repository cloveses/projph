import os
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.units import mm
from reportlab.graphics.barcode import code39, code128, code93
from ph_models import *

# 暂未使用

## 尺寸mm
ID_SIZE = (210,99)
## 水印文本
WATERMARK_TXT = "泗县教体局"
# 标题及其位置
TITLE = (50*mm,84*mm,"2018年初中学业水平考试")
ID_NAME = (55*mm,74*mm,"准　考　证")
## mm
# POSITIONS = ((5,12),(5,9),(5,6),(5,3),(12,8))
POS_X = 30
POS_Y = (64,54,44,34,24,14)
## 以上为以下七项的输出位置
ITEM_NAMES = ("准考号:  ","姓　名:  ","性　别:  ","考　点:  ","报名点:  ")

STUDS = [('183222501402',"李文娟",'男','泗县一中','01中学','tt.png'),
        ('183222501409',"张文地",'男','泗县一中','01中学','tt.png'),
        ('183222501503',"小文地",'男','泗县二中','01中学','tt.png')]
IMG_PATH = ".\\idsd"

BAR_METHODS = {'code39':code39.Extended39, 
            'code128':code128.Code128,
            'code93':code93.Standard93}

# 条形码打印位置
BAR_X = 120
BAR_Y = 14

# 照片打印位置
IMG_X = 120
IMG_Y = 24

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
    canv.drawString(BAR_X*mm+42*mm,BAR_Y*mm,"‡")
    # barcode39 = code39.Extended39('34322545666',barHeight=1*cm,barWidth=0.8)
    # barcode39.drawOn(c,20,20)
    # barcode93 = code93.Standard93('34322545666')
    # barcode93.drawOn(c,20,60)
    # barcode128 = code128.Code128('34322545666')
    # barcode128.drawOn(c,20,100)

def draw_page(canv,stud):
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
    canv.drawImage(stud[-1],IMG_X*mm,IMG_Y*mm)
    # 绘制条形码
    draw_barcode(canv,stud[0])
    # 绘制水印
    set_font(canv,10)
    canv.setFillColorRGB(180,180,180,alpha=0.3)
    canv.drawString(IMG_X*mm+mm,IMG_Y*mm+0.5*mm,WATERMARK_TXT)
    canv.showPage()

def gen_pdf(dir_name,sch_name,studs):
    confirm_path(dir_name)
    path = os.path.join(dir_name,sch_name + '.pdf')
    canv = canvas.Canvas(path,pagesize=(ID_SIZE[0]*mm,ID_SIZE[1]*mm))
    for stud in studs:
        draw_page(canv,stud)
    canv.save()

# def gen_all_pdfs(dir_name):
#     schs = getdata.get_schs()
#     for sch in schs:
#         studs = getdata.get_studs(sch)
#         gen_pdf(dir_name,sch,studs)


#以下各函数中照片文件未生成

# # 按学校生成准考证
# def gen_examid_sch(dir_name):
#     schs = select(s.sch for s in StudPh)
#     for sch in schs:
#         datas = select((s.phid,s.name,s.sex,s.exam_addr,s.sch,''.join(("Z",s.signid,'.jpg')))
#          for s in StudPh if s.sch==sch).order_by(StudPh.classcode,StudPh.phid)
#         gen_pdf(dir_name,sch,datas)

# # # 按时间段（半日）和考点生成准考证
# def gen_bak_examid(dir_name):
#     exam_addrs = select(s.exam_addr for s in StudPh)
#     exam_dates = select(s.exam_date for s in StudPh)
#     for exam_date in exam_dates:
#         for exam_addr in exam_addrs:
#             studs = select((s.phid,s.name,s.sex,s.exam_addr,s.sch,''.join(("Z",s.signid,'.jpg')))
#                 for s in StudPh if s.exam_addr==exam_addr and s.exam_date==exam_date).order_by(StudPh.phid)
#             gen_pdf(dir_name,exam_addr+exam_date,studs)
