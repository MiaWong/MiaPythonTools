#! /usr/bin/env python
# -*- coding:utf-8 -*-

# author MiaWang

# 用来统计每天提交的bug数
import pymysql.cursors
import xlwt
import sys
import importlib
importlib.reload(sys)

#这个方法是用来统计每天新开的bug数
def bugCountNewOpenByDate(productCode, sheet_name, outputpath):
    # 创建数据连接，连接到禅道数据库
    zentaoDBConnect = pymysql.connect(host="172.17.1.200",port=3306,user='root',passwd='这里需要输入密码',db='zentao',charset='utf8')
    cur = zentaoDBConnect.cursor()

    # 创建Workbook
    workbook = xlwt.Workbook()
    # 创建表
    sheet = workbook.add_sheet(sheet_name, cell_overwrite_ok=True)

    # 筛选每个模块不同的严重程度，对应的bug数量，写入表格
    sqlcount = "SELECT cast(openedDate as date)  ,COUNT(1) from zt_bug b WHERE b.deleted = '0' and " \
               "b.module in (SELECT zt_module.id from zt_module LEFT JOIN zt_product ON zt_module.root = zt_product.id " \
               "WHERE zt_product.`code` = '%s')  GROUP BY cast(openedDate as date)  order by cast(openedDate as date) asc" %(productCode)
    re_count = cur.execute(sqlcount)
    result_count = cur.fetchall()

    for i in range (len(result_count)):
        for j in range (len(result_count[i])):
            sheet.write(i, j, '%s' % result_count[i][j])

    # 保存文件
    workbook.save(outputpath)

    cur.close()
    zentaoDBConnect.close()

#这个方法是用来统计每天bug数（新开+已有）
def bugCountOpenByDate(productCode, sheet_name, outputpath):
    # 创建数据连接，连接到禅道数据库
    zentaoDBConnect = pymysql.connect(host="172.17.1.200",port=3306,user='root',passwd='Mia123456',db='zentao',charset='utf8')
    cur = zentaoDBConnect.cursor()

    # 创建Workbook
    workbook = xlwt.Workbook()
    # 创建表
    sheet = workbook.add_sheet(sheet_name, cell_overwrite_ok=True)

    # 筛选每个模块不同的严重程度，对应的bug数量，写入表格
    sqlcount = "SELECT cast(openedDate as date)  ,COUNT(1) from zt_bug b WHERE b.deleted = '0' and " \
               "b.module in (SELECT zt_module.id from zt_module LEFT JOIN zt_product ON zt_module.root = zt_product.id " \
               "WHERE zt_product.`code` = '%s')  GROUP BY cast(openedDate as date)  order by cast(openedDate as date) asc" %(productCode)
    re_count = cur.execute(sqlcount)
    result_count = cur.fetchall()

    bugcount = 0
    for i in range (len(result_count)):
        for j in range (len(result_count[i])):
            if(j == 1):
                newopenbugcount = result_count[i][j]
                bugcount += newopenbugcount
                sheet.write(i, j, '%s' % bugcount)
            else:
                sheet.write(i, j, '%s' % result_count[i][j])

    # 保存文件
    workbook.save(outputpath)

    cur.close()
    zentaoDBConnect.close()

if __name__=='__main__':
    # 产品名
    productCode = 'NHRS'
    bugCountOpenByDate(productCode, 'Bug Count by date', 'ZentaoBugs.xlsx')

















