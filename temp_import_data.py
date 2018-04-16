from ph_models import *
import xlrd

# 修正错误临时用
@db_session
def edit_phid_temp():
    wb = xlrd.open_workbook('data.xls')
    ws = wb.sheets()[0]
    nrows = ws.nrows
    for i in range(1,nrows):
        row = ws.row_values(i)
        stud = select(s for s in StudPh if s.signid==row[0]).first()
        # stud.phid = row[1]
        # stud.exam_addr = row[7]
        # stud.exam_date = row[8]
        stud.classcode = str(int(row[1]))

if __name__ == '__main__':
    edit_phid_temp()