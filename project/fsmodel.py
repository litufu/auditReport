# -*- coding: UTF-8 -*-

import pandas as pd

from project.utils import handleName, searchRecordItemByName, searchModel


# 检查是否存在model.xlxs中的报表项目和展示的报表项目无法匹配的项目
def testSubjectAllIn(comparativeTable, path):
    for item in comparativeTable:
        df = pd.read_excel(path, sheet_name=item["table"])
        dc = df.to_dict("split")
        dcNames = [handleName(d[0]) for d in dc["data"]]
        for record in item["model"]:
            if record["fillNum"]:
                if handleName(record["name"]) in dcNames:
                    continue
                print(record)


# 填充资产负债表，利润表和现金流量表,所有者权益变动表
def fillData(companyType, fsType, comparativeTable, path, tables, contrastSubjects):
    for table in tables:
        # 从model中读取报表
        df = pd.read_excel(path, sheet_name=table)
        dc = df.to_dict("split")
        model = searchModel(companyType, fsType, table, comparativeTable)
        if model is None:
            raise Exception("表格名称必须为：{}".format(tables))
        for item in dc["data"]:
            #TODO:与财务费用利息收入相区分
            name = handleName(item[0])
            record = searchRecordItemByName(name, model, fillNum=True)
            # if companyType == "上市公司":
            if record is None:
                if name in contrastSubjects:
                    record = searchRecordItemByName(contrastSubjects[name], model, fillNum=True)
            if not record is None:
                if "所有者权益变动表" in table:
                    record["paidInCapital"] = item[1]
                    record["preferedStock"] = item[2]
                    record["perpetualDebt"] = item[3]
                    record["otherEquityInstruments"] = item[4]
                    record["capitalReserve"] = item[5]
                    record["treasuryStock"] = item[6]
                    record["otherComprehensiveIncome"] = item[7]
                    record["specialReserve"] = item[8]
                    record["surplusReserve"] = item[9]
                    record["generalRiskReserve"] = item[10]
                    record["undistributedProfit"] = item[11]
                    record["subtotal"] = item[12]
                    record["minorityInterests"] = item[13]
                    record["totalOwnerEquity"] = item[14]
                else:
                    record["endDate"] = item[1]
                    record["startDate"] = item[2]

# 填充报表数据
def fillTable(context,comparativeTable,tables,contrastSubjects,combinePath,singlePath):
    # 公司类型
    companyType = context["report_params"]["companyType"]
    # 合并报表还是单体报表
    reportType = context["report_params"]["type"]
    # 根据model.xlsx填充项目数据
    if reportType == "合并":
        fillData(companyType, "合并", comparativeTable, combinePath, tables, contrastSubjects)
        fillData(companyType, "单体", comparativeTable, singlePath, tables, contrastSubjects)
    else:
        fillData(companyType, "单体", comparativeTable, singlePath, tables, contrastSubjects)

def test():
    from project.data import testcontext
    from project.constants import comparativeTable, MODELPATH, tables, contrastSubjects

    # 公司类型
    companyType = testcontext["report_params"]["companyType"]
    # 报表类型:合并，单体
    fsType = "合并"
    # 测试报表项目差
    # testSubjectAllIn(comparativeTable,MODELPATH)
    #     根据报表填充项目
    fillData(companyType, fsType, comparativeTable, MODELPATH, tables, contrastSubjects)
    for item in comparativeTable:
        print(item["model"])

if __name__ == '__main__':
    test()
