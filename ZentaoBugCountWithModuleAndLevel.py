#! /usr/bin/env python
# -*- coding:utf-8 -*-

# author MiaWang

# bugCountByModuleAndLevel方法 用来统计每个模块每个严重级别分别有多少个bug，也就是我们集成测试报告中的第一个统计表格
import pymysql.cursors
import xlwt
import sys
import importlib
importlib.reload(sys)

# 参数说明：
# 用来统计每个模块每个严重级别分别有多少个bug，也就是我们集成测试报告中的第一个统计表格
# 想用来统计自己测试的模块的bug，可以在调用时将产品名替换成自己的产品名称
def bugCountByModuleAndLevel(productCode, sheet_name, outputpath):
    # 创建数据连接，连接到禅道数据库
    zentaoDBConnect = pymysql.connect(host="172.17.1.200",port=3306,user='root',passwd='这里需要输入密码',db='zentao',charset='utf8')
    cur = zentaoDBConnect.cursor()

    # 查询当前项目的bug列表中，存在的模块名称
    sqlmodule = "select id, CONCAT_WS(' / ', (select p.name from zt_module p where p.id = m.parent), m.name) as modulenamewithparent from" \
                " zt_module m WHERE m.id in ( SELECT module FROM zt_bug WHERE product = ( SELECT p.id FROM zt_product p WHERE p.code = '%s'))"%(productCode)
    re_module = cur.execute(sqlmodule)
    module_list = cur.fetchall()

    # 获取所有的bug严重程度
    sqlseverity = "SELECT distinct severity FROM zt_bug"
    re_severity = cur.execute(sqlseverity)
    result_severity = cur.fetchall()

    # 创建Excel
    workbook = xlwt.Workbook()
    # 创建sheet
    sheet = workbook.add_sheet(sheet_name, cell_overwrite_ok=True)

    # 模块名称，Excel写入表格里面
    for row in range(0, 1):
        for col in range(1, len(result_severity) + 1):
            sheet.write(row, col, '%s' % result_severity[col - 1][0])
            print(result_severity[col - 1][0])

    # bug严重程度，Excel写入表格里面
    for row in range(1, len(module_list) + 1):
        for col in range(0, 1):
            sheet.write(row, col, '%s' % module_list[row - 1][1])
            print(module_list[row - 1][1])

    # 筛选每个模块不同的严重程度，对应的bug数量，写入表格
    for i in range(0, len(module_list)):
        for j in range(0, len(result_severity)):
            severity = str(result_severity[j][0])
            module = str(module_list[i][0])
            sqlcount = "SELECT count(*) from zt_bug b WHERE  b.severity = '%s' and b.deleted = '0' " \
                       "and b.module in (SELECT zt_module.id from zt_module LEFT JOIN zt_product ON zt_module.root = zt_product.id " \
                       "WHERE zt_product.`code` = '%s' and zt_module.id = '%s')" % (severity, productCode, module)
            re_count = cur.execute(sqlcount)
            result_count = cur.fetchall()
            sheet.write(i + 1, j + 1, '%s' % result_count[0])

    # 保存文件
    workbook.save(outputpath)

    cur.close()
    zentaoDBConnect.close()

if __name__=='__main__':
    # 产品名
    productCode = 'NHRS'
    bugCountByModuleAndLevel(productCode, 'bugs', 'ZentaoBugs.xlsx')