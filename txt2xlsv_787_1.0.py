# coding=utf-8
# code by sakura_xp 11:56 2010-10-31
# Revised by Zhushiqi from YZR 15:40 2017-8-24
# python + pyExcelerator


import pyExcelerator, sys, re
import time

# from pyExcelerator import *

######################################################################

alignC = pyExcelerator.Alignment()
alignC.horz = pyExcelerator.Alignment.HORZ_CENTER
alignC.vert = pyExcelerator.Alignment.VERT_CENTER

alignL = pyExcelerator.Alignment()
alignL.horz = pyExcelerator.Alignment.HORZ_LEFT
alignL.vert = pyExcelerator.Alignment.VERT_CENTER

alignR = pyExcelerator.Alignment()
alignR.horz = pyExcelerator.Alignment.HORZ_RIGHT
alignR.vert = pyExcelerator.Alignment.VERT_CENTER

dic_Alignment = {'left': alignL, 'center': alignC, 'right': alignR}
dic_Font = {}  # 字体
dic_Border = {}  # 表格边框
dic_Style = {}  # 单元格样式
dic_RowStyle = {}  # 行高


# 获取行高
def GetRowStyle(height):
    global dic_RowStyle
    if not dic_RowStyle.has_key(height):
        fnt = pyExcelerator.Font()
        fnt.height = height * 15
        rowStyle = pyExcelerator.XFStyle()
        rowStyle.font = fnt
        dic_RowStyle[height] = rowStyle
    return dic_RowStyle[height]


# 获取字体样式
def GetFont(fontname, bold, italic, fontsize):
    global dic_Font
    if not dic_Font.has_key((fontname, bold, italic, fontsize)):
        font = pyExcelerator.Font()
        font.name = fontname
        font.bold = bold
        font.italic = italic
        font.height = fontsize * 20
        dic_Font[(fontname, bold, italic, fontsize)] = font
    return dic_Font[(fontname, bold, italic, fontsize)]


def GetAlignment(algn):
    global dic_Alignment
    return dic_Alignment[algn]


def GetBorder(top, bottom, left, right):
    global dic_Border
    if not dic_Border.has_key((top, bottom, left, right)):
        borders0 = pyExcelerator.Borders()
        borders0.left = left
        borders0.right = right
        borders0.top = top
        borders0.bottom = bottom
        dic_Border[(top, bottom, left, right)] = borders0
    return dic_Border[(top, bottom, left, right)]


def GetStyle(top, bottom, left, right, bold=False, italic=False, fontname='Arial', fontsize=10, algn='center'):
    font = GetFont(fontname, bold, italic, fontsize)
    borders = GetBorder(top, bottom, left, right)
    alignment = GetAlignment(algn)
    global dic_Style
    if not dic_Style.has_key((font, borders, alignment)):
        style = pyExcelerator.XFStyle()
        style.font = font
        style.borders = borders
        style.alignment = alignment
        dic_Style[(font, borders, alignment)] = style
    return dic_Style[(font, borders, alignment)]


def GetStyleSimSun9pt(top, bottom, left, right, bold=False, italic=False, algn='center'):
    return GetStyle(top, bottom, left, right, bold, italic, u'宋体', 9, algn)


def GetStyleSimSun8pt(top, bottom, left, right, bold=False, italic=False, algn='center'):
    return GetStyle(top, bottom, left, right, bold, italic, u'宋体', 8, algn)


def GetStyleArial9pt(top, bottom, left, right, bold=False, italic=False, algn='center'):
    return GetStyle(top, bottom, left, right, bold, italic, 'Arial', 9, algn)


def GetStyleArial8pt(top, bottom, left, right, bold=False, italic=False, algn='center'):
    return GetStyle(top, bottom, left, right, bold, italic, 'Arial', 8, algn)


def txt2para(txtfile):
    '对源文件进行分段，并判断是否需要处理'
    return [para for para in open(txtfile).read().split("El")]

def is_not_empty(s):
    '结合filter函数去除list中的空集'
    return s and len(s.strip()) > 0


def SetColumnWidth(ws):
    ws.col(0).width = 1190
    ws.col(1).width = 1373
    ws.col(2).width = 3600
    ws.col(3).width = 3600
    ws.col(4).width = 3600
    ws.col(5).width = 3600


############################################################################
class Para:
    def Parser(self, para):

        # 处理 Head
        header_re = "^\s*(evation:.*?)\s*Runway\s*(\S+)\s*(\S+)\s*(\S+)\s*(\S+)\s*(\S+)\s*Dated (.*)\s*(Derate = \S+)\s*Flaps = Flaps\s*(\S+)" \
                    "\s*A/I = (\S+)\s*CG = (\S+)\s*Runway Cond = (\S+)\s*(QNH = \S+ mbar)\s*(V1 Policy = \S+)\s*(Time Limit = \S+ Minutes)" \
                    "\s*(Path = Second or Extended)\s*Reversers = (\S+\s*\S+)\s*(Brakes = \S+\s*\S+)\s*(Gear = \S+)" \
                    "\s*(A/S = \S+)" \
                    "\s*(PMS = \S+)" \
                    "\s*\n.*\n.*\n.*\n.*\xa1\xe3C\s*(\S+)\s*\((\S+ \S+)\)" \
                    "\s*(\S+)\s*(\S+)\s*(\S+)\s*(\S+)" \
                    "\s*\n.*"
        m = re.search(header_re, para, re.M)
        if m == None:
            sys.exit('Header Match Failed.')
        self.Elevation, self.Runway, self.AIRPORT, self.Engine_Type, self.Airport_Adress, self.Airport_City, self.Date, self.Derate, self.Flaps, self.A_I, self.CG, self.Runway_Cond, self.QNH, self.V1_Policy, \
        self.Time_Limit, self.Path, self.Reversers, self.Brakes, self.Gear, self.AS, self.PMS, self.Line, self.Climb, self.Wind1, self.Wind2, self.Wind3, self.Wind4 = m.groups()
        self.Airport_Adress_City = self.Airport_Adress + "/" +self.Airport_City
        # 整理，格式化数据项
        # 去除787机场分析中重量和速度之间的空格（如有）
        #处理BODY
        para = para[m.end():].replace(" /", "/").replace("|", "")
        lines = para.lstrip().split('\n')
        self.BODY = []
        global n # n为body的总行数
        n = 0
        for line in lines:
            if (len(line) == 0) or (line[0:3] == 'Min'):
                break
            n += 1
            i = 0
            if len(line.split()) < 4 and len(line.split()) > 0:
                line1 = line.split()
                line2 = []
                if line[26] == ' ':
                    line2.append(' ')
                else:
                    line2.append(line1[i])
                    i += 1
                if line[46] == ' ':
                    line2.append(' ')
                else:
                    line2.append(line1[i])
                    i += 1
                if line[66] == ' ':
                    line2.append(' ')
                else:
                    line2.append(line1[i])
                    i += 1
                if line[86] == ' ':
                    line2.append(' ')
                else:
                    line2.append(line1[i])
                    i += 1
            else:
                line2 = line.split()
            self.BODY.append(line2)
        # 处理 foot
        para = '\n'.join(lines[n:]).lstrip()
        lines = para.split('\n')
        self.FOOT = []
        FOOT_LINE_NUMBER = 0
        for line in lines:
            if line[-4:] == 'FT/M':
                break
            self.FOOT.append(line)
            FOOT_LINE_NUMBER += 1
        end = ' '.join(lines[FOOT_LINE_NUMBER+2:]).strip()
        end = end[:end.index('procedure')]
        items = end.split()

        if items < 2:
            sys.exit('HT/DIST/OFFSET Count Error!')

        items = items[1:len(items) - 2]

        if items[0] == 'No':
            pass  # 没有HT/DIST/OFFSET的情况
        elif len(items) % 3 != 0:
            pass
        # sys.exit('HT/DIST/OFFSET Count Error2!')


        if items[0] == 'No':
            self.HT_DISTS = ['%s/%s' % (HT, DIST) for HT, DIST, OFFSET in
                        [items[i * 3:i * 3 + 3] for i in range((len(items) - 1) / 3)]]
        else:
            self.HT_DISTS = ['%s/%s' % (HT, DIST) for HT, DIST, OFFSET in
                        [items[i * 3:i * 3 + 3] for i in range((len(items) + 2) / 3)]]


    def Print(self):
        pass

    def SetRowHeight(self, row):
        for i in range(0, 4): ws.row(row + i).set_style(GetRowStyle(4))
        for i in range(4, 7): ws.row(row + i).set_style(GetRowStyle(10))
        for i in range(7, 9): ws.row(row + i).set_style(GetRowStyle(11))
        for i in range(9, 41): ws.row(row + i).set_style(GetRowStyle(12.5))
        for i in range(41, 52): ws.row(row + i).set_style(GetRowStyle(9))
        for i in range(52, 54): ws.row(row + i).set_style(GetRowStyle(9))

    ##########################################################################################

    def WriteSheet_737(self, ws, row):

        ws.write_merge(row, row + 1, 0, 2, 'SUPARNA AIRLINES', GetStyle(1, 0, 1, 1, True, True))

        ws.write_merge(row + 2, row + 3, 0, 2, 'RAM B'+ self.Engine_Type[0:5], GetStyle(0, 1, 1, 1, True))
        ws.write_merge(row, row + 3, 3, 3, self.AIRPORT, GetStyle(1, 1, 1, 1, True))

        ws.write_merge(row, row + 3, 4, 4, 'RWY ' + self.Runway, GetStyle(1, 1, 1, 1, True))
        ws.write_merge(row, row + 3, 5, 5, Rev_No, GetStyle(1, 1, 1, 1, True, fontname=u'宋体', algn='right'))

        row += 4  # row 4
        ws.write_merge(row, row, 0, 2, self.Airport_Adress_City, GetStyleSimSun9pt(1, 0, 1, 0, algn='left'))
        ws.write_merge(row, row, 5, 5, 'FLAPS  ' + self.Flaps, GetStyleSimSun9pt(1, 0, 0, 1, True, algn='right'))

        ws.write_merge(row, row, 3, 3, self.Engine_Type[6:], GetStyle(1, 1, 1, 1, True, fontsize=9))
        ws.write_merge(row+1, row+1, 3, 3, self.Derate, GetStyle(1, 1, 1, 1, True, fontsize=9))
        ws.write_merge(row, row + 1, 4, 4, self.Runway_Cond + ' Runway', GetStyle(1, 1, 1, 1, True, fontsize=9))

        row += 1  # row 5

        ws.write_merge(row, row, 0, 2, 'El'+self.Elevation, GetStyleSimSun9pt(0, 1, 1, 1, algn='left'))

        ws.write_merge(row, row, 5, 5, 'RVS '+ self.Reversers, GetStyleSimSun9pt(0, 0, 1, 1, True, algn='right'))
        row += 1  # row 6
        ws.write_merge(row, row, 0, 4, '*A* INDICATES OAT OUTSIDE ENVIRONMENTAL ENVELOPE',
                       GetStyleSimSun9pt(0, 1, 1, 0, algn='left'))
        ws.write(row, 5, 'ANTI-ICE ' + self.A_I, GetStyleSimSun9pt(0, 1, 1, 1, True, algn='right'))
        row += 1  # row 7
        ws.write(row, 0, 'OAT', GetStyleSimSun9pt(1, 0, 1, 1))
        ws.write(row, 1, 'CLIMB', GetStyleSimSun9pt(1, 0, 1, 1))
        ws.write_merge(row, row, 2, 5, 'WIND COMPONENT IN KNOTS (MINUS DENOTES TAILWIND)',
                       GetStyleSimSun9pt(1, 1, 1, 1))
        row += 1  # row 8
        ws.write(row, 0, u'℃', GetStyleSimSun9pt(0, 1, 1, 1))
        ws.write(row, 1, self.Climb, GetStyleArial9pt(0, 1, 1, 1))
        ws.write(row, 2, self.Wind1, GetStyleArial9pt(1, 1, 1, 1))
        ws.write(row, 3, self.Wind2, GetStyleArial9pt(1, 1, 1, 1))
        ws.write(row, 4, self.Wind3, GetStyleArial9pt(1, 1, 1, 1))
        ws.write(row, 5, self.Wind4, GetStyleArial9pt(1, 1, 1, 1))

        # output body
        row += 1  # row 9
        row2 = 0
        z = 0
        empty = [' ', ' ', ' ', ' ', ' ', ' ']
        if self.Runway_Cond == 'Dry':
            for line in self.BODY:
                top = row2 % 6 == 0
                column = 0
                column1 = 0
                if len(line) < 6:
                    line = [''] * 2+ line[0:]
                for item in line:
                    ws.write(row + row2, column, item, GetStyleArial9pt(top, 0, 1, 1))
                    column += 1
                if z != n-1 :
                    if len(self.BODY[z]) == 6 & len(self.BODY[z+1]) == 6:
                        row2 += 1
                        for item in empty:
                            top = row2 % 6 == 0
                            ws.write(row + row2, column1, item, GetStyleArial9pt(top, 0, 1, 1))
                            column1 += 1
                elif len(self.BODY[z]) == 6:
                        row2 += 1
                        for item in empty:
                            top = row2 % 6 == 0
                            ws.write(row + row2, column1, item, GetStyleArial9pt(top, 0, 1, 1))
                            column1 += 1

                row2 += 1
                z += 1
        else:
            for line in self.BODY:
                top = row2 % 4 == 0
                column = 0
                if len(line) < 6:
                    line = line[:2] + [''] * (6 - len(line)) + line[2:]
                for item in line:
                    ws.write(row + row2, column, item, GetStyleArial9pt(top, 0, 1, 1))
                    column += 1
                row2 += 1
                #        assert row2 == n+2, 'Error'

        # output foot\


        row += row2

        self.FOOT = filter(is_not_empty,self.FOOT)
        for line in self.FOOT:
            ws.write_merge(row, row, 0, 5, line, GetStyleSimSun8pt(1, 1, 1, 1, algn='left'))
            row += 1



        l = 0
        for item in ('OBS  FROM', 'HT/DIST', 'FT/M'):
            ws.write_merge(row + l, row + l, 0, 1, item, GetStyleArial9pt(1, 1, 1, 1))
            l += 1

        if len(self.HT_DISTS) == 0:
            ws.write(row, 2, "NONE", GetStyleArial9pt(1, 1, 1, 1))
            i = 1
            while i < 12:
                row2 = i / 4
                col2 = i % 4
                ws.write(row + row2, 2 + col2, '', GetStyleArial8pt(1, 1, 1, 1))
                i += 1

        else:
            i = 0
            for item in self.HT_DISTS:
                row2 = i / 4
                col2 = i % 4
                ws.write(row + row2, 2 + col2, item, GetStyleArial8pt(1, 1, 1, 1))
                i += 1
            while i < 12:
                row2 = i / 4
                col2 = i % 4
                ws.write(row + row2, 2 + col2, '', GetStyleArial8pt(1, 1, 1, 1))
                i += 1

        row += 3

        ws.write_merge(row, row, 0, 5, u'ENGINE OUT PROCEDURE: NO SPECIAL PROCEDURE',
                       GetStyleSimSun9pt(1, 1, 1, 1, algn='left'))
        row += 1
        ws.write_merge(row, row, 0, 3, u'技术支援中心 Flight Operation Engineering',
                       GetStyleSimSun9pt(1, 1, 1, 0, algn='left'))
        ws.write_merge(row, row, 4, 5, self.Date, GetStyleSimSun9pt(1, 1, 0, 1, algn='right'))
        row += 1
        return row


def getRunway(para):

    Find_Runway = "^\s*(evation:.*?)\s*Runway\s*(\S+)\s*(\S+)\s*(\S+)\s*(\S+)\s*(\S+)\s*Dated (.*)\s*(Derate = \S+)\s*Flaps = Flaps\s*(\S+)"\
                 "\s*A/I = (\S+)\s*CG = (\S+)\s*Runway Cond = (\S+)"
    n = re.search(Find_Runway, para, re.M)
    Runway1 = n.group(2)
    Airport1 = n.group(3)
    Runway_Cond = n.group(12)
    return Runway1+Airport1+Runway_Cond
    sys.exit('GetRunway Error!')



if __name__ == '__main__':

    print '现在开始转换机场分析格式....\n请输入版本号'

    Rev_No = raw_input()
    Rev_No = 'Rev ' + Rev_No

    fname = 'input3.TXT'
    paras = txt2para(fname)
    paras = paras[1:]

    w = pyExcelerator.Workbook()
    key_value_runway_para = [(getRunway(para), para) for para in paras]
    runways = list([runway for runway, para in key_value_runway_para])

    for runway in runways:
        print runway
        ws = w.add_sheet(runway.replace('/','-'))
        row = 0
        paras_filtered = [para for runway2, para in key_value_runway_para if runway2 == runway]
        for para in paras_filtered:
            p = Para()
            p.Parser(para)
            p.Print()
            row_NextPage = p.WriteSheet_737(ws, row)
            p.SetRowHeight(row)
            row = row_NextPage
        SetColumnWidth(ws)
        # break
    w.save('787output.xls')

    start = time.clock()
    end = time.clock()
    print "Operation time is : %f s" % (end - start)
