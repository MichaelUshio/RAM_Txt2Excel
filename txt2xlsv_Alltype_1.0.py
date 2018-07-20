# coding=utf-8
# code by sakura_xp 11:56 2010-10-31
# Revised by Zhushiqi from YZR 15:40 2017-8-24
# python + pyExcelerator
# 2018-3-8 revised '/' error from Hainan Airlines to '-'
# V1.0 Edition published to GitHub ON 2018-07-20

import pyExcelerator, sys, re

# from pyExcelerator import *

######################################################################
dic_to_derate = {
    '15% DERATE': '       TO2         15% DERATE',
    '20% DERATE': '       TO2         20% DERATE',
    '10% DERATE': '       TO1         10% DERATE',
    '05% DERATE': '       TO1         05% DERATE',
}





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
    return [para for para in open(txtfile).read().split('\x0c') if para.strip().startswith('ELEVATION')]


def SetColumnWidth(ws):
    ws.col(0).width = 1190
    ws.col(1).width = 1373
    ws.col(2).width = 3495
    ws.col(3).width = 3225
    ws.col(4).width = 3410
    ws.col(5).width = 3620


############################################################################
class Para:
    def Parser(self, para):

        # 处理 Head
        header_re = "^\s*(ELEVATION.*?)\s*RUNWAY\s*(\S+)\s*(\S+)\s*\*\*\* FLAPS (\S+) \*\*\*\s*(AIR COND \S+)\s*(ANTI-ICE \S+)\s*(.*\n.*?)\s*$\s*(\S+)\s*(\S+)\s*(.*?)DATED (.*)\n.*\n.*\n\s*C\s*(\S+)\s*(\S+)\s*(\S+)\s*(\S+)\s*(\S+)\s*"

        # 处理一发失效程序
        self.ENG_OUT = re.findall("SPECIAL", para)
        if self.ENG_OUT:
            self.ENG_OUT = '*** SEE SPECIAL PROCEDURE ***'
        else:
            self.ENG_OUT = '*** NO EMERGENCY TURN ***'

        # 处理湿滑跑道刹车效应
        slippery_re = re.findall("SLI.*MU", para)
        wet_runway = re.findall("WET RWY", para)
        if slippery_re:
            if slippery_re[0][-5:-3] == '10':
                self.MU = slippery_re[0]
                self.GM = 'Middle'
                self.RUNWAY_COND = 'SLIPPERY RWY'
            elif slippery_re[0][-5:-3] == '20':
                self.MU = slippery_re[0]
                self.GM = 'Good'
                self.RUNWAY_COND = 'SLIPPERY RWY'
        elif wet_runway:
            self.RUNWAY_COND = 'WET RWY'
        else:
            self.RUNWAY_COND = 'DRY RWY'

        m = re.search(header_re, para, re.M)
        if m == None:
            sys.exit('Header Match Failed.')
        self.ELEVATION, self.RUNWAY, self.AIRPORT, self.FLAPS, self.AIR_COND, self.ANTI_ICE, self.ADDRESS, self.PLANE_TYPE, self.ENGINE_TYPE, self.TO_DERATE, self.DATE, self.CLIMB, self.WIND1, self.WIND2, self.WIND3, self.WIND4 = m.groups()
        # 整理，格式化数据项
        self.ADDRESS = ' '.join(self.ADDRESS.split())
        #        self.TO_DERATE = dic_to_derate.get(self.TO_DERATE.strip(), self.ENGINE_TYPE)

        # 处理 body
        para = para[m.end():]
        lines = para.lstrip().split('\n')
        self.BODY = []

        global n # n为body的总行数
        n = 0
        if self.PLANE_TYPE != '747-400'and self.PLANE_TYPE != '747-400F':
            for line in lines:
                n += 1
                i = 0
                if line[6:9] == 'MAX':
                    break
                if len(line.split()) < 4 and len(line.split()) > 0:
                    line1 = line.split()
                    line2 = []
                    if line[20] == ' ':
                        line2.append(' ')
                    else:
                        line2.append(line1[i])
                        i += 1
                    if line[37] == ' ':
                        line2.append(' ')
                    else:
                        line2.append(line1[i])
                        i += 1
                    if line[53] == ' ':
                        line2.append(' ')
                    else:
                        line2.append(line1[i])
                        i += 1
                    if line[69] == ' ':
                        line2.append(' ')
                    else:
                        line2.append(line1[i])
                        i += 1
                else:
                    line2 = line.split()
                self.BODY.append(line2)
        else:
            for line in lines:
                n += 1
                i = 0
                if line[6:9] == 'MAX':
                    break
            self.BODY = [line.split() for line in lines[:n-1]]

        # 处理 foot
        para = '\n'.join(lines[n-1:]).lstrip()
        lines = para.split('\n')
        self.FOOT = ('' + lines[0].strip(),
                     '' + lines[1].strip(),
                     '' + lines[2].strip(),
                     '' + lines[3].strip(),
                     '' + lines[4].strip(),
                     '' + lines[5].strip(),
                     '' + lines[6].strip(),
                     )
        end = ' '.join(lines[8:]).strip()
        end = end[:end.index('ENG-OUT')]
        items = end.split()
        if items < 2:
            sys.exit('HT/DIST/OFFSET Count Error!')
        items = items[1:]
        if items[0] == 'NONE':
            pass  # 没有HT/DIST/OFFSET的情况
        elif len(items) % 3 != 0:
            sys.exit('HT/DIST/OFFSET Count Error2!')

        if items[0] == 'NONE':
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
        for i in range(9, 43): ws.row(row + i).set_style(GetRowStyle(12.5))
        for i in range(43, 51): ws.row(row + i).set_style(GetRowStyle(9))
        ws.row(row + 51).set_style(GetRowStyle(11.25))

    ##########################################################################################
    def WriteSheet_747(self, ws, row):

        ws.write_merge(row, row + 1, 0, 2, 'SUPARNA AIRLINES', GetStyle(1, 0, 1, 1, True, True))
        if self.ENGINE_TYPE == ("CF6-80C2B1F"):
            self.PLANE_TYPE = '747-400SF'
        elif self.ENGINE_TYPE == ("CF6-80C2B5F"):
            self.PLANE_TYPE = '747-400ERF'
        else:
            sys.exit('NOT 747!')

        ws.write_merge(row + 2, row + 3, 0, 2, 'RAM B' + self.PLANE_TYPE, GetStyle(0, 1, 1, 1, True))
        ws.write_merge(row, row + 3, 3, 3, self.AIRPORT, GetStyle(1, 1, 1, 1, True))
        to = dic_to_derate.get(self.TO_DERATE.strip(), 'TO')

        ws.write_merge(row, row + 3, 4, 4, 'RWY ' + self.RUNWAY, GetStyle(1, 1, 1, 1, True))
        ws.write_merge(row, row + 1, 5, 5, Rev_No, GetStyle(1, 1, 1, 1, True, fontname=u'宋体', algn='right'))
        ws.write_merge(row + 2, row + 3, 5, 5, self.ENGINE_TYPE,
                       GetStyle(1, 1, 1, 1, True, fontname=u'宋体', algn='right'))

        row += 4  # row 4
        ws.write_merge(row, row, 0, 2, self.ADDRESS, GetStyleSimSun9pt(1, 0, 1, 0, algn='left'))
        ws.write_merge(row, row, 5, 5, 'FLAPS  ' + self.FLAPS, GetStyleSimSun9pt(1, 0, 0, 1, True, algn='right'))

        if self.RUNWAY_COND == 'DRY RWY':
            ws.write_merge(row, row, 3, 4, dic_to_derate.get(self.TO_DERATE.strip(), '       TO         NO DERATE'),
                           GetStyle(0, 0, 0, 0, True ))
        elif self.RUNWAY_COND == 'WET RWY':
            ws.write_merge(row, row, 3, 4, dic_to_derate.get(self.TO_DERATE.strip(), '       TO         NO DERATE'),
                           GetStyle(0, 0, 0, 0, True ))
        else:
            ws.write(row, 3,  'NO DERATE', GetStyle(0, 0, 0, 0, True, fontsize = 9, algn='center'))
            ws.write(row, 4, 'SLIPPERY RWY', GetStyle(0, 0, 0, 0, True , fontsize = 9, algn='center'))

        row += 1  # row 5
        ws.write_merge(row, row, 0, 2, self.ELEVATION, GetStyleSimSun9pt(0, 0, 1, 0, algn='left'))

        if self.RUNWAY_COND == 'DRY RWY':
            ws.write_merge(row, row, 3, 4, self.RUNWAY_COND, GetStyle(0, 0, 0, 0, True, fontsize=10))
        elif self.RUNWAY_COND == 'WET RWY':
            ws.write_merge(row, row, 3, 4, self.RUNWAY_COND, GetStyle(0, 0, 0, 0, True, fontsize=10))
        else:
            ws.write_merge(row, row, 3, 4, 'Brake action is ' + self.GM + '(0' + self.MU[12:] + ')',
                           GetStyle(0, 0, 0, 0, True, fontsize=8))

        ws.write_merge(row, row, 5, 5, self.AIR_COND, GetStyleSimSun9pt(0, 0, 0, 1, True, algn='right'))
        row += 1  # row 6
        ws.write_merge(row, row, 0, 4, '*A* INDICATES OAT OUTSIDE ENVIRONMENTAL ENVELOPE',
                       GetStyleSimSun9pt(0, 1, 1, 0, algn='left'))
        ws.write(row, 5, self.ANTI_ICE, GetStyleSimSun9pt(0, 1, 0, 1, True, algn='right'))
        row += 1  # row 7
        ws.write(row, 0, 'OAT', GetStyleSimSun9pt(1, 0, 1, 1))
        ws.write(row, 1, 'CLIMB', GetStyleSimSun9pt(1, 0, 1, 1))
        ws.write_merge(row, row, 2, 5, 'WIND COMPONENT IN KNOTS (MINUS DENOTES TAILWIND)',
                       GetStyleSimSun9pt(1, 1, 1, 1))
        row += 1  # row 8
        ws.write(row, 0, u'℃', GetStyleSimSun9pt(0, 1, 1, 1))
        ws.write(row, 1, self.CLIMB, GetStyleArial9pt(0, 1, 1, 1))
        ws.write(row, 2, self.WIND1, GetStyleArial9pt(1, 1, 1, 1))
        ws.write(row, 3, self.WIND2, GetStyleArial9pt(1, 1, 1, 1))
        ws.write(row, 4, self.WIND3, GetStyleArial9pt(1, 1, 1, 1))
        ws.write(row, 5, self.WIND4, GetStyleArial9pt(1, 1, 1, 1))

        # output body
        row += 1  # row 9
        row2 = 0
        for line in self.BODY:
            top = row2 % 5 == 0
            column = 0
            if len(line) < 6:
                line = line[:2] + [''] * (6 - len(line)) + line[2:]
            for item in line:
                ws.write(row + row2, column, item, GetStyleArial9pt(top, 0, 1, 1))
                column += 1
            row2 += 1
#        assert row2 == 33, 'Error'

        # output foot
        row += row2-1  # row 42
        for line in self.FOOT:
            ws.write_merge(row, row, 0, 5, line, GetStyleSimSun8pt(1, 1, 1, 1, algn='left'))
            row += 1

        l = 0
        for item in ('OBS  FROM', 'HT/DIST', 'FT/FT'):
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

        row += 3  # row 51
        # ENGINE OUT PROCEDURE
        ws.write_merge(row, row, 0, 5, u'ENGINE OUT PROCEDURE:' + self.ENG_OUT,
                       GetStyleSimSun9pt(1, 1, 1, 1, algn='left'))
        row += 1
        ws.write_merge(row, row, 0, 3, u'技术支援中心 Flight Operation Engineering',
                       GetStyleSimSun9pt(1, 1, 1, 0, algn='left'))
        ws.write_merge(row, row, 4, 5, self.DATE, GetStyleSimSun9pt(1, 1, 0, 1, algn='right'))
        row += 1
        return row

    def WriteSheet_737(self, ws, row):

        ws.write_merge(row, row + 1, 0, 2, 'SUPARNA AIRLINES', GetStyle(1, 0, 1, 1, True, True))
        if SFP == 1:
            self.PLANE_TYPE = '737-800W-26K-SFP1'
        elif SFP == 2:
            self.PLANE_TYPE = '737-800W-26K-SFP2'
        elif self.ENGINE_TYPE == ("CFM56-7B26"):
            self.PLANE_TYPE = '737-800W-26K'
        elif self.ENGINE_TYPE == ("CFM56-7B24"):
            self.PLANE_TYPE = '737-800W-24K'
        elif self.ENGINE_TYPE == ("CFM56-3B-2"):
            self.PLANE_TYPE = '737-400-22K'
        elif self.ENGINE_TYPE == ("CFM56-3-B1"):
            self.PLANE_TYPE = '737-300-20K'
        else:
            sys.exit('NOT 737!')

        ws.write_merge(row + 2, row + 3, 0, 2, 'RAM B' + self.PLANE_TYPE, GetStyle(0, 1, 1, 1, True))
        ws.write_merge(row, row + 3, 3, 3, self.AIRPORT, GetStyle(1, 1, 1, 1, True))
        to = dic_to_derate.get(self.TO_DERATE.strip(), 'TO')

        ws.write_merge(row, row + 3, 4, 4, 'RWY ' + self.RUNWAY, GetStyle(1, 1, 1, 1, True))
        ws.write_merge(row, row + 3, 5, 5, Rev_No, GetStyle(1, 1, 1, 1, True, fontname=u'宋体', algn='right'))

        row += 4  # row 4
        ws.write_merge(row, row, 0, 2, self.ADDRESS, GetStyleSimSun9pt(1, 0, 1, 0, algn='left'))
        ws.write_merge(row, row, 5, 5, 'FLAPS  ' + self.FLAPS, GetStyleSimSun9pt(1, 0, 0, 1, True, algn='right'))

        ws.write_merge(row, row + 1, 3, 3, self.ENGINE_TYPE, GetStyle(1, 1, 1, 1, True, fontsize=9))
        ws.write_merge(row, row + 1, 4, 4, self.RUNWAY_COND, GetStyle(1, 1, 1, 1, True, fontsize=9))

        row += 1  # row 5

        ws.write_merge(row, row, 0, 2, self.ELEVATION, GetStyleSimSun9pt(0, 1, 1, 1, algn='left'))



        ws.write_merge(row, row, 5, 5, self.AIR_COND, GetStyleSimSun9pt(0, 0, 1, 1, True, algn='right'))
        row += 1  # row 6
        ws.write_merge(row, row, 0, 4, '*A* INDICATES OAT OUTSIDE ENVIRONMENTAL ENVELOPE',
                       GetStyleSimSun9pt(0, 1, 1, 0, algn='left'))
        ws.write(row, 5, self.ANTI_ICE, GetStyleSimSun9pt(0, 1, 1, 1, True, algn='right'))
        row += 1  # row 7
        ws.write(row, 0, 'OAT', GetStyleSimSun9pt(1, 0, 1, 1))
        ws.write(row, 1, 'CLIMB', GetStyleSimSun9pt(1, 0, 1, 1))
        ws.write_merge(row, row, 2, 5, 'WIND COMPONENT IN KNOTS (MINUS DENOTES TAILWIND)',
                       GetStyleSimSun9pt(1, 1, 1, 1))
        row += 1  # row 8
        ws.write(row, 0, u'℃', GetStyleSimSun9pt(0, 1, 1, 1))
        ws.write(row, 1, self.CLIMB, GetStyleArial9pt(0, 1, 1, 1))
        ws.write(row, 2, self.WIND1, GetStyleArial9pt(1, 1, 1, 1))
        ws.write(row, 3, self.WIND2, GetStyleArial9pt(1, 1, 1, 1))
        ws.write(row, 4, self.WIND3, GetStyleArial9pt(1, 1, 1, 1))
        ws.write(row, 5, self.WIND4, GetStyleArial9pt(1, 1, 1, 1))

        # output body
        row += 1  # row 9
        row2 = 0
        z = 0
        empty = [' ', ' ', ' ', ' ', ' ', ' ']
        if self.RUNWAY_COND == 'DRY RWY':
            for line in self.BODY:
                top = row2 % 6 == 0
                column = 0
                column1 = 0
                if len(line) < 6:
                    line = [''] * 2+ line[0:]
                for item in line:
                    ws.write(row + row2, column, item, GetStyleArial9pt(top, 0, 1, 1))
                    column += 1
                if z != n - 2:
                    if len(self.BODY[z + 1]) == 6 & len(self.BODY[z]) == 6:
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

        # output foot

        if len(self.BODY[z - 2]) == 6:
            if self.RUNWAY_COND == 'DRY RWY':
                row += row2
            else:
                row += row2 - 1
        else:
            row += row2 - 1

        for line in self.FOOT:
            ws.write_merge(row, row, 0, 5, line, GetStyleSimSun8pt(1, 1, 1, 1, algn='left'))
            row += 1

        l = 0
        for item in ('OBS  FROM', 'HT/DIST', 'M/M'):
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

        # ENGINE OUT PROCEDURE
        ws.write_merge(row, row, 0, 5, u'ENGINE OUT PROCEDURE:' + self.ENG_OUT,
                       GetStyleSimSun9pt(1, 1, 1, 1, algn='left'))
        row += 1
        ws.write_merge(row, row, 0, 3, u'技术支援中心 Flight Operation Engineering',
                       GetStyleSimSun9pt(1, 1, 1, 0, algn='left'))
        ws.write_merge(row, row, 4, 5, self.DATE, GetStyleSimSun9pt(1, 1, 0, 1, algn='right'))
        row += 1
        return row


def getRunway(para):
    start1 = para.find('RUNWAY ')
    start2 = 74
    end2 = 78
    if start1 >= 0:
        start1 += len('RUNWAY ')
        end1 = para.find(' ', start1)
        if end1 >= 0:
            Runway = para[start2: end2] + '(' + para[start1: end1] + ')' + para[114:120].replace('ON','AUTO') + para[38:46]
            return Runway
            # else

    sys.exit('GetRunway Error!')



if __name__ == '__main__':

    print '现在开始转换机场分析格式....\n请输入版本号'

    Rev_No = raw_input()
    Rev_No = 'Rev ' + Rev_No


    fname = 'input.TXT'
    paras = txt2para(fname)
#判断机型
    AC_TYPE = re.findall("7.7-...", paras[0])
    SFP = 0
    if AC_TYPE[0][4:8] == '800':
        print '请确认机型为非SFP,SFP1,SFP2:\n1-SFP1\n2-SFP2\n0-非SFP'
        SFP = input()
#判断是不是747 AUTO 三个 DERATE
    DERATE_INFO = re.findall("..% DERATE", paras[0])
    w = pyExcelerator.Workbook()
    key_value_runway_para = [(getRunway(para), para) for para in paras]

#如果要输出DERATE，就必须用SET函数，否则使用LIST，才可以输出重复的跑道
    
    if DERATE_INFO:
        runways = sorted(set([runway for runway, para in key_value_runway_para]))
    else:
        runways = list([runway for runway, para in key_value_runway_para])

    for runway in runways:
        print runway
        ws = w.add_sheet(runway.replace('/','-').replace('ON','auto'))
        row = 0
        paras_filtered = [para for runway2, para in key_value_runway_para if runway2 == runway]
        for para in paras_filtered:
            p = Para()
            p.Parser(para)
            p.Print()
            if AC_TYPE[0] == '747-400':
                row_NextPage = p.WriteSheet_747(ws, row)
            else:
                row_NextPage = p.WriteSheet_737(ws, row)
            p.SetRowHeight(row)
            row = row_NextPage
        SetColumnWidth(ws)
        # break
    w.save('output.xls')
