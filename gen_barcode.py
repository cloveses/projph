import os
import math
from ph_models import *

from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.units import mm
from reportlab.graphics.barcode import code39, code128, code93

BAR_METHODS = {'code39':code39.Extended39, 
            'code128':code128.Code128,
            'code93':code93.Standard93}

def get_positions():
    poses = []
    base_x = 15
    base_y = 22
    col_xes = (i * 50 + base_x for i in range(4))
    for col_x in col_xes:
        col_poses = [(col_x,i * base_y + base_x) for i in range(12)]
        poses.extend(col_poses)
    return poses



def set_font(canv,size,font_name='simsun',font_file='simsun.ttc'):
    pdfmetrics.registerFont(TTFont(font_name,font_file))
    canv.setFont(font_name,size)

def draw_barcode(canv,idcode,pos,codetype='code128'):
    ## 绘制条形码
    ## codetype have:code39,code93,code128
    barcd = BAR_METHODS[codetype](idcode,barWidth=1,humanReadable=True)
    barcd.drawOn(canv,pos[0]*mm,pos[1]*mm)
    set_font(canv,16)
    canv.drawString(pos[0]*mm+10,pos[1]*mm,"‡")
    canv.drawString(pos[0]*mm+34.2*mm,pos[1]*mm,"‡")
    # barcode39 = code39.Extended39('34322545666',barHeight=1*cm,barWidth=0.8)
    # barcode39.drawOn(c,20,20)
    # barcode93 = code93.Standard93('34322545666')
    # barcode93.drawOn(c,20,60)
    # barcode128 = code128.Code128('34322545666')
    # barcode128.drawOn(c,20,100)

phids = [i for i in range(18250001,18250049)]

def draw_page(canv,sch_name,phids):
    # 绘制标题
    set_font(canv,11)
    canv.drawString(80*mm,280*mm,sch_name)
    for pos,phid in zip(get_positions(),phids):
        draw_barcode(canv,phid,pos)

    canv.showPage()

def gen_pdf(sch_name,phids):
    pagenum = 48
    path = ''.join((sch_name,'.pdf'))
    canv = canvas.Canvas(path)
    for i in range(math.ceil(len(phids)/pagenum)):
        draw_page(canv,sch_name,phids[i*pagenum:(i+1)*pagenum])
    canv.save()

# 以学校为单位生成条形码文件
@db_session
def get_sch_barcodes():
    schs = select(s.sch for s in StudPh)
    for sch in schs:
        studs = select(s for s in StudPh if s.sch==sch).order_by(StudPh.sex).order_by(StudPh.phid)
        phids = [s.phid for s in studs]
        gen_pdf(sch,phids)

if __name__ == '__main__':
    get_sch_barcodes()