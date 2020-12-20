# -*- coding: UTF-8 -*-

from project.utils import searchRecordItemByName,searchModel

# 计算数字编码
# 计算资产项目编码
def computeAssetsNum(records,startNum):
    for key, item in enumerate(records):
        if item["hasNum"] and (abs(item["startDate"]) > 1e-6 or abs(item["endDate"]) > 1e-6):
            records[key]["noteNum"] = startNum
            startNum += 1
        else:
            records[key]["noteNum"] = ""
    return records,startNum
# 计算负债项目编码
def computeLiabilitiesNum(liabilitiesRecords,assetsRecords,startNum,):
    for key, item in enumerate(liabilitiesRecords):
        if item["name"] == "递延所得税负债":
            if abs(item["startDate"]) > 1e-6 or abs(item["endDate"]) > 1e-6:
                deferTaxAsset = searchRecordItemByName("递延所得税资产", assetsRecords)
                if deferTaxAsset["noteNum"] != "":
                    liabilitiesRecords[key]["noteNum"] = deferTaxAsset["noteNum"]
                else:
                    liabilitiesRecords[key]["noteNum"] = startNum
                    startNum += 1
            else:
                liabilitiesRecords[key]["noteNum"] = ""
        else:
            if item["hasNum"] and (abs(item["startDate"]) > 1e-6 or abs(item["endDate"]) > 1e-6):
                liabilitiesRecords[key]["noteNum"] = startNum
                startNum += 1
            else:
                liabilitiesRecords[key]["noteNum"] = ""
    return liabilitiesRecords,startNum
# 计算利润表编码
def computeProfitNum(profitRecords,liabilitiesRecords,startNum,ociInProfitName,incomeName,costName):
    for key, item in enumerate(profitRecords):
        if item["name"] == ociInProfitName:
            if abs(item["startDate"]) > 1e-6 or abs(item["endDate"]) > 1e-6:
                oci = searchRecordItemByName("其他综合收益", liabilitiesRecords)
                if oci["noteNum"] != "":
                    profitRecords[key]["noteNum"] = oci["noteNum"]
                else:
                    profitRecords[key]["noteNum"] = startNum
                    startNum += 1
            else:
                profitRecords[key]["noteNum"] = ""
        elif item["name"] == costName:
            if abs(item["startDate"]) > 1e-6 or abs(item["endDate"]) > 1e-6:
                bussinessIncome = searchRecordItemByName(incomeName, profitRecords)
                if bussinessIncome["noteNum"] != "":
                    profitRecords[key]["noteNum"] = bussinessIncome["noteNum"]
                else:
                    profitRecords[key]["noteNum"] = startNum
                    startNum += 1
            else:
                profitRecords[key]["noteNum"] = ""
        else:
            if item["hasNum"] :
                if abs(item["startDate"]) > 1e-6 or abs(item["endDate"]) > 1e-6:
                    profitRecords[key]["noteNum"] = startNum
                    startNum += 1
                else:
                    profitRecords[key]["noteNum"] = ""
            else:
                profitRecords[key]["noteNum"] = ""
    return profitRecords,startNum
# 计算现金流量表编码
def computeCashNum(cashRecords,startNum):
    for key, item in enumerate(cashRecords):
        if item["hasNum"] and (abs(item["startDate"]) > 1e-6 or abs(item["endDate"]) > 1e-6):
            cashRecords[key]["noteNum"] = startNum
            startNum += 1
        else:
            cashRecords[key]["noteNum"] = ""
    return cashRecords,startNum
# 计算单体资产表编码
def computeSingleAssetNum(recordsSingle,startNumSingle,names):
    for key, item in enumerate(recordsSingle):
        if item["name"] in names and (abs(item["startDate"]) > 1e-6 or abs(item["endDate"]) > 1e-6):
            recordsSingle[key]["noteNum"] = startNumSingle
            startNumSingle += 1
        else:
            recordsSingle[key]["noteNum"] = ""
    return recordsSingle,startNumSingle
# 为报表中所有项目添加空附注
def computeNoneNoteNum(records):
    for key, item in enumerate(records):
        records[key]["noteNum"] = ""
    return records
# 计算单体利润表编码
def computeSingleProfitNum(profitRecordsSingle,startNumSingle,incomeName,costName):
    for key, item in enumerate(profitRecordsSingle):
        if item["name"] == incomeName and (abs(item["startDate"]) > 1e-6 or abs(item["endDate"]) > 1e-6):
            profitRecordsSingle[key]["noteNum"] = startNumSingle
            startNumSingle += 1
        elif item["name"] == costName:
            if abs(item["startDate"]) > 1e-6 or abs(item["endDate"]) > 1e-6:
                bussinessIncome = searchRecordItemByName(incomeName, profitRecordsSingle)
                if bussinessIncome["noteNum"] != "":
                    profitRecordsSingle[key]["noteNum"] = bussinessIncome["noteNum"]
                else:
                    profitRecordsSingle[key]["noteNum"] = startNumSingle
                    startNumSingle += 1
            else:
                profitRecordsSingle[key]["noteNum"] = ""
        elif item["name"] == "投资收益（损失以“-”号填列）" and (abs(item["startDate"]) > 1e-6 or abs(item["endDate"]) > 1e-6):
            profitRecordsSingle[key]["noteNum"] = startNumSingle
            startNumSingle += 1
        else:
            profitRecordsSingle[key]["noteNum"] = ""
    return profitRecordsSingle,startNumSingle

# 计算国有企业报表数据编码
# 逻辑关系：
# 按照公司类型分分别上市公司和国企
# 上市公司报表下分合并报告和单体报告
#     上市公司合并报表和单体报表格式不相同
# 国企分合并报告和单体报告
#     国企合并报表和单体报表格式相同
# 普通：
#     如果该项目属于有编码的项目，则从1开始递增。
# 编码公用：
#     递延所得税资产和递延所得税负债共用一个编码
#     营业收入和营业成本公用一个编码
#     利润表其他综合收益与资产负债表其他综合收益公用一个编码
# 如果是合并报告，单体报表只有个别项目需要标注：应收账款、其他应收款、长期股权投资、营业收入和营业成本、研发费用、投资收益

def computeNoteNum(assetsRecordsCombine,liabilitiesRecordsCombine,profitRecordsCombine,cashRecordsCombine,
                   assetsRecordsSingle,liabilitiesRecordsSingle,profitRecordsSingle,cashRecordsSingle,context):

    startNum = 1
    startNumSingle = 1
    companyType = context["report_params"]["companyType"]
    reportType = context["report_params"]["type"]
    if  companyType == "上市公司":
        # 上市公司合并报告
        if reportType =="合并":
            # 添加合并资产表附注编号
            assetsRecordsCombine,startNum = computeAssetsNum(assetsRecordsCombine, startNum)
            # 添加合并负债表附注编号
            liabilitiesRecordsCombine,startNum = computeLiabilitiesNum(liabilitiesRecordsCombine,assetsRecordsCombine,startNum)
            # 添加合并利润表附注编号
            profitRecordsCombine,startNum = computeProfitNum(profitRecordsCombine, liabilitiesRecordsCombine, startNum, "六、其他综合收益的税后净额", "其中：营业收入", "其中：营业成本")
            # 添加合并现金流量表编号
            cashRecordsCombine,startNum = computeCashNum(cashRecordsCombine,startNum)
            # 添加单体资产表编号
            assetsRecordsSingle,startNumSingle = computeSingleAssetNum(assetsRecordsSingle, startNumSingle,["应收账款","其他应收款","长期股权投资"])
            # 添加单体负债表编号
            computeNoneNoteNum(liabilitiesRecordsSingle)
            # 添加单体利润表编号
            profitRecordsSingle, startNumSingle = computeSingleProfitNum(profitRecordsSingle, startNumSingle, "一、营业收入", "减：营业成本")
             # 添加单体现金流量表编号
            computeNoneNoteNum(cashRecordsSingle)
        # 上市公司单体报告
        else:
            # 添加单体资产表附注编号
            assetsRecordsSingle, startNum = computeAssetsNum(assetsRecordsSingle, startNum)
            #  添加单体负债表附注编号
            liabilitiesRecordsSingle, startNum = computeLiabilitiesNum(liabilitiesRecordsSingle, assetsRecordsSingle,
                                                                        startNum)
            # 添加利润表附注编号
            profitRecordsSingle, startNum = computeProfitNum(profitRecordsSingle, liabilitiesRecordsSingle, startNum,
                                                              "五、其他综合收益的税后净额", "一、营业收入", "减：营业成本")
            # 添加现金流量表编号
            cashRecordsSingle,startNum = computeCashNum(cashRecordsSingle,startNum)
    elif companyType == "国有企业":
        # 国有企业合并报告
        if reportType == "合并":
            # 添加合并资产表附注编号
            assetsRecordsCombine, startNum = computeAssetsNum(assetsRecordsCombine, startNum)
            # 添加合并负债表附注编号
            liabilitiesRecordsCombine, startNum = computeLiabilitiesNum(liabilitiesRecordsCombine, assetsRecordsCombine,
                                                                        startNum)
            # 添加合并利润表附注编号
            profitRecordsCombine, startNum = computeProfitNum(profitRecordsCombine, liabilitiesRecordsCombine, startNum,
                                                              "六、其他综合收益的税后净额", "其中：营业收入", "其中：营业成本")
            # 添加合并现金流量表
            computeNoneNoteNum(cashRecordsCombine)
            # 添加单体资产表编号
            assetsRecordsSingle, startNumSingle = computeSingleAssetNum(assetsRecordsSingle, startNumSingle,
                                                                        ["应收账款", "其他应收款", "长期股权投资"])
            # 添加单体负债表编号
            computeNoneNoteNum(liabilitiesRecordsSingle)
            # 添加单体利润表编号
            profitRecordsSingle, startNumSingle = computeSingleProfitNum(profitRecordsSingle, startNumSingle, "其中：营业收入",
                                                                         "其中：营业成本")
            #添加单体现金流
            computeNoneNoteNum(cashRecordsSingle)
        # 国有企业单体报告
        else:
            # 添加单体资产表附注编号
            assetsRecordsSingle, startNum = computeAssetsNum(assetsRecordsSingle, startNum)
            #  添加单体负债表附注编号
            liabilitiesRecordsSingle, startNum = computeLiabilitiesNum(liabilitiesRecordsSingle, assetsRecordsSingle,
                                                                       startNum)
            # 添加利润表附注编号
            profitRecordsSingle, startNum = computeProfitNum(profitRecordsSingle, liabilitiesRecordsSingle, startNum,
                                                             "六、其他综合收益的税后净额", "其中：营业收入", "其中：营业成本")
            #    添加单体现金流量表
            computeNoneNoteNum(cashRecordsSingle)

def computeNo(context,comparativeTable):
    # 公司类型
    companyType = context["report_params"]["companyType"]
    # 获取报表数据
    assetsRecordsCombine = searchModel(companyType, "合并", "资产表", comparativeTable)
    liabilitiesRecordsCombine = searchModel(companyType, "合并", "负债表", comparativeTable)
    profitRecordsCombine = searchModel(companyType, "合并", "利润表", comparativeTable)
    cashRecordsCombine = searchModel(companyType, "合并", "现金流量表", comparativeTable)
    assetsRecordsSingle = searchModel(companyType, "单体", "资产表", comparativeTable)
    liabilitiesRecordsSingle = searchModel(companyType, "单体", "负债表", comparativeTable)
    profitRecordsSingle = searchModel(companyType, "单体", "利润表", comparativeTable)
    cashRecordsSingle = searchModel(companyType, "单体", "现金流量表", comparativeTable)

    # 计算附注编码
    computeNoteNum(assetsRecordsCombine, liabilitiesRecordsCombine, profitRecordsCombine, cashRecordsCombine,
                   assetsRecordsSingle, liabilitiesRecordsSingle, profitRecordsSingle, cashRecordsSingle, context)
