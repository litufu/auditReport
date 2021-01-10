# -*- coding: UTF-8 -*-

import os
from project.data import initData
from project.report import auditReport
from project.financialStatements import addFs
from project.notes import addNoteAccountingPolicy
from project.changeAndErrorCorrection import addChange
from project.tax import addTax
from project.combine import addCombine
from project.noteappended import addNoteAppended
from project.fsmodel import fillTable
from project.computeNo import computeNo
from project.constants import comparativeTable,tables,contrastSubjects,balanceTitlesState,balanceTitlesList,profitTitlesState,profitTitlesList


if __name__ == '__main__':
    CURRENTPATH = "D:/我的文件2021/义务2020年审/义乌市粮食收储有限公司TB及附注/1义乌市粮油质量检测中心有限公司.xlsx"
    PARENTPATH = "D:/我的文件2021/义务2020年审/义乌市粮食收储有限公司TB及附注/1义乌市粮油质量检测中心有限公司.xlsx"
    parent_path = os.path.dirname(CURRENTPATH)
    filename = os.path.basename(CURRENTPATH)
    new_filename = "{}.docx".format(filename.replace(".xlsx","").replace("TB",""))
    new_path = os.path.join(parent_path,new_filename)
    # 数据初始化
    context = initData(CURRENTPATH)
    # 公司类型：上市公司、国有企业
    companyType = context["report_params"]["companyType"]
    # 报告类型：合并、单体
    reportType = context["report_params"]["type"]
    # 添加报告
    document = auditReport(context)
    # 填充报表数据
    fillTable(context, comparativeTable, tables, contrastSubjects, CURRENTPATH, PARENTPATH)
    # 计算附注编码
    computeNo(context, comparativeTable)
    # 添加报表
    addFs(document, context, comparativeTable, balanceTitlesState, balanceTitlesList, profitTitlesState,profitTitlesList)
    # 添加会计政策,返回政策编码
    num = addNoteAccountingPolicy(document,context,comparativeTable)
    # 添加会计政策变更
    addChange(document, num, context, CURRENTPATH,PARENTPATH)
    # 添加税收政策
    addTax(document, context, CURRENTPATH)
    # 国有企业添加企业合并及合并财务报表
    if companyType == "国有企业" and reportType == "合并":
        addCombine(document, CURRENTPATH, context)
    # 添加财务报表注释
    addNoteAppended(document, CURRENTPATH,PARENTPATH,context, comparativeTable, isAll=False)


    document.save(new_path)
