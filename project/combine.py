# -*- coding: UTF-8 -*-

from docx import Document
import pandas as pd

from project.data import report_params,combine
from project.settings import setStyle
from project.utils import addTitle, addParagraph,addTable,addLandscapeContent,createBorderedTable,addContentToCombineTitle,addCombineTableTitle

# （一）	子企业情况
def addContent1(document,reportType,startYear,MODELPATH):
    addTitle(document, "七、企业合并及合并财务报表", 1, False)
    addTitle(document, "（一）子企业情况", 2, True)
    df = pd.read_excel(MODELPATH, sheet_name="子企业情况")
    dc = df.to_dict("split")
    addTable(document, dc, style=6)

# （二）母公司拥有被投资单位表决权不足半数但能对被投资单位形成控制的原因
def addContent2(document,reportType,startYear,MODELPATH):
    addTitle(document, "（二）母公司拥有被投资单位表决权不足半数但能对被投资单位形成控制的原因", 2, True)
    df = pd.read_excel(MODELPATH, sheet_name="表决权不足半数但能形成控制")
    dc = df.to_dict("split")
    addTable(document, dc, style=6)

# （三）母公司直接或通过其他子公司间接拥有被投资单位半数以上表决权但未能对其形成控制的原因
def addContent3(document,reportType,startYear,MODELPATH):
    addTitle(document, "（三）母公司直接或通过其他子公司间接拥有被投资单位半数以上表决权但未能对其形成控制的原因", 2, True)
    df = pd.read_excel(MODELPATH, sheet_name="半数以上表决权但未控制")
    dc = df.to_dict("split")
    addTable(document, dc, style=6)

# （四）重要非全资子企业情况
def addContent4(document,reportType,startYear,MODELPATH):
    addTitle(document, "（四）重要非全资子企业情况", 2, True)
    addParagraph(document,"1、少数股东","paragraph")
    df = pd.read_excel(MODELPATH, sheet_name="重要非全资子企业少数股东")
    dc = df.to_dict("split")
    addTable(document, dc, style=6)
    addParagraph(document,"2、主要财务信息","paragraph")
    # 资产负债信息
    addParagraph(document, "（1）资产和负债情况", "paragraph")
    df = pd.read_excel(MODELPATH, sheet_name="重要非全资子企业期末资产负债")
    dc = df.to_dict("split")
    titles = [["单位名称", "期末数", "nan", "nan", "nan", "nan", "nan"],
              ["nan", "流动资产", "非流动资产", "资产合计", "流动负债", "非流动负债", "负债合计"]]
    titleLength = len(titles)
    rowLength = len(dc["index"]) + titleLength
    columnLength = len(dc["columns"])
    table = createBorderedTable(document, rowLength, columnLength)
    addCombineTableTitle(table, titles)
    addContentToCombineTitle(document, dc, table, titleLength, style=3)
    addParagraph(document, "(续上表)：".format(startYear), "paragraph")
    df = pd.read_excel(MODELPATH, sheet_name="重要非全资子企业期初资产负债")
    dc = df.to_dict("split")
    titles = [["单位名称", "期初数", "nan", "nan", "nan", "nan", "nan"],
              ["nan", "流动资产", "非流动资产", "资产合计", "流动负债", "非流动负债", "负债合计"]]
    titleLength = len(titles)
    rowLength = len(dc["index"]) + titleLength
    columnLength = len(dc["columns"])
    table = createBorderedTable(document, rowLength, columnLength)
    addCombineTableTitle(table, titles)
    addContentToCombineTitle(document, dc, table, titleLength, style=3)

    # 损益及现金流信息
    addParagraph(document, "（2）损益和现金流量情况", "paragraph")
    df = pd.read_excel(MODELPATH, sheet_name="重要非全资子企业本期损益和现金流量情况")
    dc = df.to_dict("split")
    titles = [["单位名称", "本期数", "nan", "nan", "nan"], ["nan", "营业收入", "净利润", "综合收益总额", "经营活动现金流量"]]
    titleLength = len(titles)
    rowLength = len(dc["index"]) + titleLength
    columnLength = len(dc["columns"])
    table = createBorderedTable(document, rowLength, columnLength)
    addCombineTableTitle(table, titles)
    addContentToCombineTitle(document, dc, table, titleLength, style=3)
    addParagraph(document, "(续上表)：".format(startYear), "paragraph")
    df = pd.read_excel(MODELPATH, sheet_name="重要非全资子企业上期损益和现金流量情况")
    dc = df.to_dict("split")
    titles = [["单位名称", "上期数", "nan", "nan", "nan"], ["nan", "营业收入", "净利润", "综合收益总额", "经营活动现金流量"]]
    titleLength = len(titles)
    rowLength = len(dc["index"]) + titleLength
    columnLength = len(dc["columns"])
    table = createBorderedTable(document, rowLength, columnLength)
    addCombineTableTitle(table, titles)
    addContentToCombineTitle(document, dc, table, titleLength, style=3)


# （五）子公司与母公司会计期间不一致的说明
def addContent5(document,reportType,startYear,MODELPATH):
    addTitle(document, "（五）子公司与母公司会计期间不一致的说明", 2, True)
    addParagraph(document,combine["statementOnInconsistencyOfAccountingPeriodBetweenSubsidiaryCompanyAndParentCompany"],"paragraph")
# （六）本年不再纳入合并范围的原子公司
def addContent6(document,reportType,startYear,MODELPATH):
    addTitle(document, "（六）本年不再纳入合并范围的原子公司", 2, True)
    addParagraph(document,"1、本年不再纳入合并范围原子公司的情况","paragraph")
    df = pd.read_excel(MODELPATH, sheet_name="本年不再纳入合并范围原子公司的情况")
    dc = df.to_dict("split")
    addTable(document, dc, style=6)
    addParagraph(document,"2、原子公司在处置日和上一会计期间资产负债表日的财务状况","paragraph")
    df = pd.read_excel(MODELPATH, sheet_name="原子公司在处置日和上一会计期间资产负债表日的财务状况")
    dc = df.to_dict("split")
    titles = [["原子公司名称","处置日","处置日","nan","nan","上期末","nan","nan"],["nan","nan","资产总额","负债总额","所有者权益总额","资产总额","负债总额","所有者权益总额"]]
    titleLength = len(titles)
    rowLength = len(dc["index"]) + titleLength
    columnLength = len(dc["columns"])
    table = createBorderedTable(document,rowLength,columnLength)
    addCombineTableTitle(table,titles)
    addContentToCombineTitle(document, dc, table, titleLength, style=3)
    addParagraph(document,"3、原子公司本年年初至处置日的经营成果","paragraph")
    df = pd.read_excel(MODELPATH, sheet_name="原子公司本年年初至处置日的经营成果")
    dc = df.to_dict("split")
    titles = [["原子公司名称","处置日","本年初至处置日","nan","nan"],["nan","nan","收入","费用","净利润"]]
    titleLength = len(titles)
    rowLength = len(dc["index"]) + titleLength
    columnLength = len(dc["columns"])
    table = createBorderedTable(document,rowLength,columnLength)
    addCombineTableTitle(table,titles)
    addContentToCombineTitle(document, dc, table, titleLength, style=3)
# （七）本年新纳入合并范围的主体
def addContent7(document,reportType,startYear,MODELPATH):
    addTitle(document, "（七）本年新纳入合并范围的主体", 2, True)
    df = pd.read_excel(MODELPATH, sheet_name="本年新纳入合并范围的主体")
    dc = df.to_dict("split")
    addTable(document, dc, style=6)
# （八）本年发生的同一控制下企业合并情况
def addContent8(document,reportType,startYear,MODELPATH):
    addTitle(document, "（八）本年发生的同一控制下企业合并情况", 2, True)
    df = pd.read_excel(MODELPATH, sheet_name="本年发生的同一控制下企业合并情况")
    dc = df.to_dict("split")
    titles = [["公司名称","合并日","合并日确定依据","账面净资产","交易对价","实际控制人","本年初至合并日的相关情况","nan","nan","nan"],["nan","nan","nan","nan","nan","nan","收入","净利润","现金净增加额","经营活动现金流量净额"]]
    titleLength = len(titles)
    rowLength = len(dc["index"]) + titleLength
    columnLength = len(dc["columns"])
    table = createBorderedTable(document,rowLength,columnLength)
    addCombineTableTitle(table,titles)
    addContentToCombineTitle(document, dc, table, titleLength, style=3)
# （九）本年发生的非同一控制下企业合并情况
def addContent9(document,reportType,startYear,MODELPATH):
    addTitle(document, "（九）本年发生的非同一控制下企业合并情况", 2, True)
    df = pd.read_excel(MODELPATH, sheet_name="本年发生的非同一控制下企业合并情况")
    dc = df.to_dict("split")
    addTable(document, dc, style=6)
# （十）本年发生的反向购买
def addContent10(document,reportType,startYear,MODELPATH):
    addTitle(document, "（十）本年发生的反向购买", 2, True)
    df = pd.read_excel(MODELPATH, sheet_name="本年发生的反向购买")
    dc = df.to_dict("split")
    addTable(document, dc, style=6)
# （十一）本年发生的吸收合并
def addContent11(document,reportType,startYear,MODELPATH):
    addTitle(document, "（十一）本年发生的吸收合并", 2, True)
    df = pd.read_excel(MODELPATH, sheet_name="本年发生的反向购买")
    dc = df.to_dict("split")
    addTable(document, dc, style=6)
# （十二）子企业使用企业集团资产和清偿企业集团债务的重大限制
def addContent12(document,reportType,startYear,MODELPATH):
    addTitle(document, "（十二）子企业使用企业集团资产和清偿企业集团债务的重大限制", 2, True)
    addParagraph(document,combine["majorRestrictionsOnSubsidiariesUseOfEnterpriseGroupAssetsAndSettlementOfEnterpriseGroupDebts"],"paragraph")
# （十三）纳入合并财务报表范围的结构化主体的相关信息
def addContent13(document,reportType,startYear,MODELPATH):
    addTitle(document, "（十三）纳入合并财务报表范围的结构化主体的相关信息", 2, True)
    addParagraph(document,combine["structuredSubject"],"paragraph")
# （十四）母公司在子企业的所有者权益份额发生变化的情况
def addContent14(document,reportType,startYear,MODELPATH):
    addTitle(document, "（十四）母公司在子企业的所有者权益份额发生变化的情况", 2, True)
    ownerEquityChange = combine["changesInShareOfOwnerEquityOfParentCompanyInSubsidiaryEnterprises"]
    addParagraph(document,ownerEquityChange,"paragraph")
    if ownerEquityChange != "不适用":
        df = pd.read_excel(MODELPATH, sheet_name="母公司在子企业的所有者权益份额发生变化的情况")
        dc = df.to_dict("split")
        addTable(document, dc, style=6)
# （十五）子公司向母公司转移资金的能力受到严格限制的情况
def addContent15(document,reportType,startYear,MODELPATH):
    addTitle(document, "（十五）子公司向母公司转移资金的能力受到严格限制的情况", 2, True)
    addParagraph(document,combine["theAbilityOfSubsidiaryToTransferFundsToItsParentCompanyIsStrictlyRestricted"],"paragraph")

def addContent(document,reportType,startYear,MODELPATH):
    addContent1(document,reportType,startYear,MODELPATH)
    addContent2(document,reportType,startYear,MODELPATH)
    addContent3(document,reportType,startYear,MODELPATH)
    addContent4(document,reportType,startYear,MODELPATH)
    addContent5(document,reportType,startYear,MODELPATH)
    addContent6(document,reportType,startYear,MODELPATH)
    addContent7(document,reportType,startYear,MODELPATH)
    addContent8(document,reportType,startYear,MODELPATH)
    addContent9(document,reportType,startYear,MODELPATH)
    addContent10(document,reportType,startYear,MODELPATH)
    addContent11(document,reportType,startYear,MODELPATH)
    addContent12(document,reportType,startYear,MODELPATH)
    addContent13(document,reportType,startYear,MODELPATH)
    addContent14(document,reportType,startYear,MODELPATH)
    addContent15(document,reportType,startYear,MODELPATH)

def addCombine(document):
    # 报告日期
    reportDate = report_params["reportDate"]
    # 获取报告起始日
    startYear = reportDate[:4]
    # 合并报表还是单体报表
    reportType = report_params["type"]
    MODELPATH = "D:/auditReport/project/model.xlsx"
    addLandscapeContent(document, addContent, reportType, startYear, MODELPATH)

def test():
    document = Document()
    setStyle(document)
    addCombine(document)
    document.save("combine.docx")

if __name__ == '__main__':
    test()