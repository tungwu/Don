#!/usr/bin/python
# -*- coding: utf8 -*-
import MySQLdb
import datetime
import  xlwt
import sys

# Excel Style variable Begin

bg_style_5 = xlwt.easyxf(
    "pattern: pattern solid, fore_colour 5;font: colour_index black, bold off; align: wrap on, vert centre, horiz center; borders: top THIN, bottom THIN, left THIN, right THIN;")  # 背景色
bg_style_15 = xlwt.easyxf(
    "pattern: pattern solid, fore_colour 15;font: colour_index black, bold off; align: wrap on, vert centre, horiz center; borders: top THIN, bottom THIN, left THIN, right THIN;")  # 背景色
bg_style_47 = xlwt.easyxf(
    "pattern: pattern solid, fore_colour 47;font: colour_index black, bold off; align: wrap on, vert centre, horiz center; borders: top THIN, bottom THIN, left THIN, right THIN;")  # 背景色

tittle_style = xlwt.easyxf(
    'font: height 300, colour_index black, bold off; align: wrap on, vert centre, horiz center;'      "borders: top NO_LINE, bottom NO_LINE, left NO_LINE, right NO_LINE;")
bottom_style = xlwt.easyxf(
    'font: name Arial Black, colour_index red, bold on; align: wrap on, vert centre, horiz center;'      "borders: top NO_LINE, bottom NO_LINE, left NO_LINE, right NO_LINE;")
right_style = xlwt.easyxf(
    'font: name Arial Black, colour_index black; align: wrap on, vert centre, horiz center;'      "borders: top THIN, bottom THIN, left THIN, right THIN;")
each_style = xlwt.easyxf(
    'font: colour_index black, bold off; align: wrap on, vert centre, horiz center;'      "borders: top THIN, bottom THIN, left THIN, right THIN;")
fontcenter = xlwt.easyxf('align: wrap on, vert centre, horiz center;')

# Excel Style variable End

# 预定义相关数组 Excel里面通用的一些表头
szsp_title = ['首字', '首屏', '流量'] * 5
szspll_site_title = ['3G', '手腾首字', '手腾首屏','手腾流量', '网易首字', '网易首屏', '网易流量','搜狐首字', '搜狐首屏', '搜狐流量', '新浪首字', '新浪首屏', '新浪流量','3G首字', '3G首屏', '3G流量']
browser = ['uc', 'qqnew', 'qqolder']
site = ['info.3g.qq.com', '3g.163.com', 'm.sohu.com', 'sina.cn', '3g.cn']

# 整理数据的方法

# 第一个函数: 用来处理 分站点 站点分首字首屏 首字首屏数据切片 首字首屏数据排序 首字首屏数据逐条写入Excel
def cells_sort_calc(rowsz=5, colsz=0, rowsp=5, colsp=1, rowll=5, colll=2, cells_col=['A','B','C'], dingbulie = 1, dibujunzhilie = 0,x = 3,letter = 'C'):
    each_site_log = []
    for i in range(speedLoglistLength):
        for j in range(3):
            each_site_log.append(speedLoglist[i][j])
    # 切片
    shouzi = each_site_log[0::3]
    shouping = each_site_log[1::3]
    liuliang = each_site_log[2::3]
    # 排序
    shouzi.sort()
    shouping.sort()
    liuliang.sort()
    # 写入sheet
    for shouziline in shouzi:
        sheet.write(rowsz, colsz, shouziline)
        rowsz = rowsz + 1

    for cells_colline in range(len(cells_col)):
        # 每列数据的去头去尾求平均值（每三格一组 首字 首屏 流量）
        sheet.write(rowsz, dibujunzhilie,xlwt.Formula('AVERAGE(%s%s:%s%s)' % (cells_col[cells_colline], excelRowAVG(shouzi)[0], cells_col[cells_colline], excelRowAVG(shouzi)[1])),bottom_style)
        # 顶部第二行引用均值单元格的值
        sheet.write(1, dingbulie, xlwt.Formula('TEXT(SUM(%s%s/1000),"0.0000")' % (cells_col[cells_colline], (rowsz + 1))), right_style)
        sheet.write(1, x, xlwt.Formula('TEXT(SUM(%s%s),"0.00")' % (letter, (rowsz + 1))), right_style)
        dingbulie = dingbulie + 1
        dibujunzhilie = dibujunzhilie + 1
    for shoupingline in shouping:
        sheet.write(rowsp, colsp, shoupingline)
        rowsp = rowsp + 1
    for liuliangline in liuliang:
        sheet.write(rowll, colll, liuliangline)
        rowll = rowll + 1



# 第二个函数: 用来生成数据报告页面

def DateSummary(x = 0, browserone = '003主线', browsertwo = '线上包', sheet1 = 'qqnew', sheet2 = 'qqolder', sum2 = 'SUM(B6,C6,E6,F6,H6,I6,K6,L6,N6,O6)', sum1 = 'SUM(B7,C7,E7,F7,H7,I7,K7,L7,N7,O7)',
                avg1 ='AVERAGE(B3,E3,H3,K3,N3)',avg2 ='AVERAGE(B4,E4,H4,K4,N4)',avg3 ='AVERAGE(D3,G3,J3,M3,P3)', avg4 ='AVERAGE(D4,G4,J4,M4,P4)',p1 =1, p2 =2,p3 =3,p4 =4,p5 =5,p6 =6,z1 ="SUM(Q6)",z2 ="SUM(Q7)",background_cells =bg_style_5):
    sheet.write_merge(x, x, 0, 18, "%s VS %s" % (browserone, browsertwo), xlwt.easyxf('font: name Arial Black, colour_index red, bold on; align: wrap on, vert centre, horiz center;' "borders: top THIN, bottom THIN, left THIN, right THIN;"))
    for x in range(len(szspll_site_title)):
        sheet.write(p1, x, szspll_site_title[x], each_style)
    sheet.write(p1, 0, "%s" % newnetwork, each_style)
    sheet.write(p2, 0, "%s" % browserone, each_style)
    sheet.write(p3, 0, "%s" % browsertwo, each_style)
    sheet.write(p4, 0, "%s提升" % browserone, each_style)
    sheet.write(p5, 0, "%s得分" % browserone, each_style)
    sheet.write(p6, 0, "%s得分" % browsertwo, each_style)
    datarow = [chr(i).upper() for i in range(97, 123)][1:16]
    calcrow = [chr(i).upper() for i in range(97, 123)][1:16]

    data_col = 1
    for datarowline in range(len(datarow)):
        sheet.write(p2, data_col, xlwt.Formula("SUM('%s %s'!%s2)" % (newnetwork, sheet1, datarow[datarowline])), each_style)
        sheet.write(p3, data_col, xlwt.Formula( "SUM('%s %s'!%s2)" % (newnetwork, sheet2, datarow[datarowline])), each_style)
        data_col = data_col + 1

    calc_col = 1
    for calcrowline in range(len(calcrow)):
        sheet.write(p4, calc_col, xlwt.Formula("SUM((%s%s-%s%s)/%s%s)" %(calcrow[calcrowline],p4,calcrow[calcrowline],p3,calcrow[calcrowline],p4)), each_style)
        if calc_col not in (3, 6, 9, 12, 15):
            sheet.write(p5, calc_col, xlwt.Formula("SUM(IF(%s%s<-0.05,0,IF(%s%s>0.05,3,1)))" % (calcrow[calcrowline],p5, calcrow[calcrowline],p5)),each_style)
            sheet.write(p6, calc_col, xlwt.Formula("SUM(IF(%s%s=3,0,IF(%s%s=0,3,1)))" % (calcrow[calcrowline],p6,calcrow[calcrowline],p6)), each_style)
        calc_col = calc_col + 1


    sheet.write(p1, 16, '速度总分', background_cells)
    sheet.write(p5, 16, xlwt.Formula(sum2), background_cells)
    sheet.write(p6, 16, xlwt.Formula(sum1), background_cells)
    sheet.write(p2, 16, ' ', each_style)
    sheet.write(p3, 16, ' ', each_style)
    sheet.write(p4, 16, ' ', each_style)

    sheet.write(p1, 17, '平均首字', background_cells)
    sheet.write(p5, 17, ' ', each_style)
    sheet.write(p6, 17, ' ', each_style)
    sheet.write(p2, 17, xlwt.Formula(avg1), background_cells)
    sheet.write(p3, 17, xlwt.Formula(avg2), background_cells)
    sheet.write(p4, 17, ' ', each_style)

    sheet.write(p1, 18, '平均流量', background_cells)
    sheet.write(p5, 18, ' ', each_style)
    sheet.write(p6, 18, ' ', each_style)
    sheet.write(p2, 18, xlwt.Formula(avg3), background_cells)
    sheet.write(p3, 18, xlwt.Formula(avg4), background_cells)
    sheet.write(p4, 18, ' ', each_style)

    sheet.write(p1, 20, "接入点", background_cells)
    sheet.write(p1, 21, "%s得分" % browserone, background_cells)
    sheet.write(p1, 22, "%s得分" % browsertwo, background_cells)
    sheet.write(p2, 20, "%s" % newnetwork, each_style)
    sheet.write(p2, 21, xlwt.Formula(z1), each_style)
    sheet.write(p2, 22, xlwt.Formula(z2), each_style)



# 第三个函数: 调用第一个函数 第二个函数 将数据库获取的数据按一定的格式样式排列生成数据表及数据汇总表页面
def speedLog(network):
    global newnetwork
    newnetwork = network
    # 先判断当天有无数据
    Day_Data = "select count(*) from daily_performance_ios_speedtimer where network = '%s' and site <> '启动' and task_id BETWEEN %s and %s" % (network, beginDate, endDate)
    cursor.execute(Day_Data)
    count = cursor.fetchone()
    # 在数据不为空的情况下进行后续操作
    if count == (0L,):
        print('查询到今天的 (%s) 数据为空!!!' % network)
    else:
        # 两个条件: 浏览器类型字段browser 和 站点类型字段site 同时满足
        # 效果：browser循环一次，site需要循环五次, browser循环三次，site循环十五次
        for browserline in range(len(browser)):  # browser长度做为循环的次数 3次
            # 每当进入browser循环一次，就需要创建一张sheet，循环三次正好就是3张sheet
            global sheet
            sheet = ws.add_sheet(network + ' %s' % browser[browserline], cell_overwrite_ok=True)
            # 列宽 杭高
            for col in range(30):
                sheet.col(col).width = 2500  # 列宽
                for row in range(80):
                    sheet.row(row).height_mismatch = True  # 行高
                    sheet.row(row).height = 420

            # 创建的表头 合并
            sheet.write_merge(2, 2, 0, 14, '%s %s' %
                              (browser[browserline], network), tittle_style)
            sheet.write_merge(3, 3, 0, 2, 'info.3g.qq.com', fontcenter)
            sheet.write_merge(3, 3, 3, 5, '3g.163.com', fontcenter)
            sheet.write_merge(3, 3, 6, 8, 'm.sohu.com', fontcenter)
            sheet.write_merge(3, 3, 9, 11, 'sina.cn', fontcenter)
            sheet.write_merge(3, 3, 12, 14, '3g.cn', fontcenter)

            for szsp_titleline in range(len(szsp_title)):
                sheet.write(4, szsp_titleline, szsp_title[szsp_titleline], xlwt.easyxf('font: name Arial Black, colour_index red, bold off; align: wrap on, vert centre, horiz center;'))

            for siteline in range(len(site)):  # site的长度作为循环的次数，5次
                # 接下来就是操作连接数据库 写入SQL 操作数据库 操作数据 保存Excel
                szspll_result = "select first_time,first_all_time,flow from daily_performance_ios_speedtimer where network = '%s' and site <> '启动' and browser = '%s'  and site = '%s' and task_id between %s and %s" % (network, browser[browserline], site[siteline], beginDate, endDate)

                # 操作数据库
                cursor.execute(szspll_result)

                # 返回结果集
                global speedLoglist
                speedLoglist = cursor.fetchall()

                # 接收结果集的长度，可以用作判断SQL执行后有没有获取相应的结果
                global speedLoglistLength
                speedLoglistLength = len(speedLoglist)
                # if判断一下，在查到数据的情况下进行后续的操作，反之不用
                if speedLoglistLength != 0:
                    if siteline == 0:
                        cells_sort_calc(rowsz=5, colsz=0, rowsp=5, colsp=1, rowll=5, colll=2, cells_col=['A','B','C'], dingbulie=1, dibujunzhilie=0,x = 3,letter='C')
                    if siteline == 1:
                        cells_sort_calc(rowsz=5, colsz=3, rowsp=5, colsp=4, rowll=5, colll=5, cells_col=['D','E','F'], dingbulie=4, dibujunzhilie=3, x = 6,letter='F')
                    if siteline == 2:
                        cells_sort_calc(rowsz=5, colsz=6, rowsp=5, colsp=7, rowll=5, colll=8, cells_col=['G','H','I'], dingbulie=7, dibujunzhilie=6, x = 9,letter='I')
                    if siteline == 3:
                        cells_sort_calc(rowsz=5, colsz=9, rowsp=5, colsp=10, rowll=5, colll=11, cells_col=['J', 'K', 'L'], dingbulie=10, dibujunzhilie=9, x = 12,letter='L')
                    if siteline == 4:
                        cells_sort_calc(rowsz=5, colsz=12, rowsp=5, colsp=13, rowll=5, colll=14, cells_col=['M', 'N', 'O'], dingbulie=13, dibujunzhilie=12, x = 15,letter='O')

            szspll_site_titlelineline = 0
            for szspll_site_titleline in range(len(szspll_site_title)):
                sheet.write(0, szspll_site_titlelineline, szspll_site_title[szspll_site_titleline], right_style)
                szspll_site_titlelineline = szspll_site_titlelineline + 1
            sheet.write(0, 0, "%s" % network, right_style)
            sheet.write(1, 0, "%s" % browser[browserline], right_style)

        # 生成报告页面
        sheet = ws.add_sheet('%s_结果' % network, cell_overwrite_ok=True)
        # 列宽 行高
        for col in range(30):
            sheet.col(col).width = 2370  # 列宽
            for row in range(80):
                sheet.row(row).height_mismatch = True  # 行高
                sheet.row(row).height = 500
        sheet.col(20).width = 3333  # 3333 = 1" (one inch).
        sheet.col(21).width = 3333  # 3333 = 1" (one inch).
        sheet.col(22).width = 3333  # 3333 = 1" (one inch).
        # 第一个
        DateSummary(x = 0, browserone = '003主线', browsertwo = '线上包', sheet1 = 'qqnew', sheet2 = 'qqolder', sum2= 'SUM(B6,C6,E6,F6,H6,I6,K6,L6,N6,O6)', sum1='SUM(B7,C7,E7,F7,H7,I7,K7,L7,N7,O7)',
                    avg1='AVERAGE(B3,E3,H3,K3,N3)', avg2='AVERAGE(B4,E4,H4,K4,N4)', avg3 ='AVERAGE(D3,G3,J3,M3,P3)', avg4 ='AVERAGE(D4,G4,J4,M4,P4)',p1 =1, p2 =2,p3 =3,p4 =4,p5 =5,p6 =6, z1 ="SUM(Q6)", z2 = "SUM(Q7)",background_cells=bg_style_5)
        # 第二个
        DateSummary(x=8, browserone='003主线', browsertwo='UC浏览器', sheet1='qqnew', sheet2='uc',sum2='SUM(B14,C14,E14,F14,H14,I14,K14,L14,N14,O14)',sum1='SUM(B15,C15,E15,F15,H15,I15,K15,L15,N15,O15)',
                    avg1='AVERAGE(B11,E11,H11,K11,N11)', avg2='AVERAGE(B12,E12,H12,K12,N12)', avg3='AVERAGE(D11,G11,J11,M11,P11)', avg4='AVERAGE(D12,G12,J12,M12,P12)', p1=9, p2=10, p3=11, p4=12, p5=13, p6=14,z1="SUM(Q14)", z2="SUM(Q15)",background_cells=bg_style_47)
        # 第三个
        DateSummary(x=16, browserone='线上包', browsertwo='UC浏览器', sheet1='qqolder', sheet2='uc',sum2='SUM(B22,C22,E22,F22,H22,I22,K22,L22,N22,O22)', sum1='SUM(B23,C23,E23,F23,H23,I23,K23,L23,N23,O23)',
                    avg1='AVERAGE(B19,E19,H19,K19,N19)', avg2='AVERAGE(B20,E20,H20,K20,N20)', avg3='AVERAGE(D19,G19,J19,M19,P19)', avg4='AVERAGE(D20,G20,J20,M20,P20)', p1=17, p2=18, p3=19, p4=20,p5=21, p6=22,z1="SUM(Q22)", z2="SUM(Q23)",background_cells=bg_style_15)


# 第四个函数: 下载每个地方跑的测试数据生成原始数据页面
def  downloadLogForEachPlace():
    area = "select area from daily_performance_ios_speedtimer where site <> '启动' and task_id  between %s and %s" % (
        beginDate, endDate)
    cursor.execute(area)
    areaAll = cursor.fetchall()
    arealist = []
    for i in range(len(areaAll)):
        for j in range(len(areaAll[i])):
            arealist.append(areaAll[i][j])
    area_only_list = list(set(arealist))
    for arealine in range(len(area_only_list)):
        if area_only_list[arealine] == 'dl':
            sheetname = '大连'
        if area_only_list[arealine] == 'sz':
            sheetname = '深圳'
        if area_only_list[arealine] == 'wh':
            sheetname = '武汉'
        if area_only_list[arealine] == 'cd':
            sheetname = '成都'
        sheet = ws.add_sheet(sheetname)
        downallbyarea = "select * from daily_performance_ios_speedtimer where site <> '启动' and area = '%s' and task_id  between %s and %s ORDER BY site, browser,first_time" % (
            area_only_list[arealine], beginDate, endDate)
        cursor.execute(downallbyarea)
        All = cursor.fetchall()
        for x in range(len(All)):
            for y in range(len(All[x])):
                sheet.write(x, y, All[x][y])



# 第五个函数: 根据每个站点跑了多少组数据预先规定好计算平均值在Excel中单元格的范围
def excelRowAVG(Number):
    global xx, yy
    xx , yy = 0 , 0
    if len(Number) == 1: xx = 6;yy = 6
    if len(Number) == 2: xx = 6;yy = 7
    if len(Number) == 3: xx = 6;yy = 8
    if len(Number) == 4: xx = 6;yy = 9
    if len(Number) == 5: xx = 6;yy = 10
    if len(Number) == 6: xx = 6;yy = 11
    if len(Number) == 7: xx = 6;yy = 12
    if len(Number) == 8: xx = 6;yy = 13
    if len(Number) == 9: xx = 6;yy = 14
    if len(Number) == 10: xx = 6;yy = 15
    if len(Number) == 11: xx = 6;yy = 16
    if len(Number) == 12: xx = 7;yy = 16
    if len(Number) == 13: xx = 7;yy = 17
    if len(Number) == 14: xx = 8;yy = 17
    if len(Number) == 15: xx = 8;yy = 18
    if len(Number) == 16: xx = 9;yy = 18
    if len(Number) == 17: xx = 9;yy = 19
    if len(Number) == 18: xx = 10;yy = 19
    if len(Number) == 19: xx = 10;yy = 20
    if len(Number) == 20: xx = 11;yy = 20
    if len(Number) == 40: xx = 16;yy = 35
    if len(Number) == 60: xx = 21;yy = 50
    if len(Number) == 80: xx = 26;yy = 65
    if len(Number) == 100: xx = 31;yy = 80
    if len(Number) == 120: xx = 36;yy = 95
    if len(Number) == 140: xx = 41;yy = 110
    if len(Number) == 160: xx = 46;yy = 125
    if len(Number) == 180: xx = 51;yy = 140
    if len(Number) == 200: xx = 56;yy = 155
    if len(Number) == 220: xx = 61;yy = 170
    if len(Number) == 240: xx = 66;yy = 185
    if len(Number) == 260: xx = 71;yy = 200
    if len(Number) == 280: xx = 76;yy = 215
    if len(Number) == 300: xx = 81;yy = 230
    if len(Number) == 320: xx = 86;yy = 245
    if len(Number) == 340: xx = 91;yy = 260
    if len(Number) == 360: xx = 96;yy = 275
    if len(Number) == 380: xx = 101;yy = 290
    if len(Number) == 400: xx = 106;yy = 305
    if len(Number) == 420: xx = 111;yy = 320
    if len(Number) == 440: xx = 116;yy = 335
    if len(Number) == 460: xx = 121;yy = 350
    if len(Number) == 480: xx = 126;yy = 365
    if len(Number) == 500: xx = 131;yy = 380
    if len(Number) == 520: xx = 136;yy = 395
    if len(Number) == 540: xx = 141;yy = 410
    if len(Number) == 560: xx = 146;yy = 425
    return [xx,yy]


# 主路径: 连接数据库 调用xlwt库 执行各个函数
# 与数据库建立连接
conn = MySQLdb.connect(
    host='daily.oa.com', user='daily_user', port=13306, passwd='daily123456', db='daily')
cursor = conn.cursor()
print ("数据库连接成功......")

# 获取时间
datesks = raw_input('请输入开始日期：')
datesjz = raw_input('请输入截止日期：')
while True:
        #如果输入时间且只输入日期8位    开始日期的0点至截止日期的24点
    if len(datesks) == 8 and len(datesjz) == 8:
        beginDate = int('%s000000' % datesks)
        endDate = int('%s235959' % datesjz)
        pass
    elif len(datesks) == 10 and len(datesjz) == 10:
        beginDate = int('%s0000' % datesks)
        endDate = int('%s5959' % datesjz)
        pass
    elif len(datesks) == 12 and len(datesjz) == 12:
        beginDate = int('%s00' % datesks)
        endDate = int('%s59' % datesjz)
        pass
    elif len(datesks) == 14 and len(datesjz) == 14:
        beginDate = int('%s' % datesks)
        endDate = int('%s' % datesjz)
        pass
    else:
        print('---输入错误---')
        datesks = raw_input('请输入开始时间：')
        datesjz = raw_input('请输入截止时间：')
        continue
    break

# 操作Excel需要用到xlwt库 这里就开始创建一个Excel
ws = xlwt.Workbook(encoding='utf-8')
print '数据正在整理请稍候......'
# 调用方法
speedLog('3g')

speedLog('wifi')

downloadLogForEachPlace()

ws.save('/Users/wudong/Desktop/iPhoneQQ浏览器_性能每日监控数据同步_主线_%s.xls' %
        datetime.datetime.now().strftime('%Y%m%d'))
print ("速度测试数据下载完成,Excel已保存桌面!")

# 关闭游标 关闭数据库
cursor.close()
conn.close()
sys.exit()
