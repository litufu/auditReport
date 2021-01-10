# -*- coding: UTF-8 -*-

from docx import Document
import pandas as pd

from project.settings import setStyle
from project.utils import addLandscapeContent, addTitle, addCombineTableTitle, addCombineTableContent, addParagraph, \
    createBorderedTable, addTable, to_chinese


# 首次执行日前后金融资产分类和计量对比表
def addContent1(document, currentpath,parentpath, context):
    reportType = context["report_params"]["type"]
    # 报告日期
    reportDate = context["report_params"]["reportDate"]
    # 获取报告起始日
    startYear = reportDate[:4]
    companyType = context["report_params"]["companyType"]

    if companyType == "国有企业":
        addParagraph(document, "（1)首次执行日前后金融资产分类和计量对比表", "paragraph")
    else:
        addParagraph(document, "①首次执行日前后金融资产分类和计量对比表", "paragraph")
    if reportType == "合并":
        # 添加合并报表影响
        addParagraph(document, "对本集团的影响：".format(startYear), "paragraph")
        df = pd.read_excel(currentpath, sheet_name="首次执行日前后金融资产分类和计量对比表")
        dc = df.to_dict("split")
        titles = [["原金融工具准则", "nan", "nan", "新金融工具准则 ", "nan", "nan"], ["项目", "计量类别", "账面价值", "项目", "计量类别", "账面价值"]]
        titleLength = len(titles)
        rowLength = len(dc["index"]) + titleLength
        columnLength = len(dc["columns"])
        table = createBorderedTable(document, rowLength, columnLength)
        addCombineTableTitle(table, titles)
        addCombineTableContent(table, dc, titleLength)
        #     添加母公司影响
        addParagraph(document, "对本公司的影响：".format(startYear), "paragraph")
        df = pd.read_excel(parentpath, sheet_name="首次执行日前后金融资产分类和计量对比表")
        dc = df.to_dict("split")
        titles = [["原金融工具准则", "nan", "nan", "新金融工具准则 ", "nan", "nan"], ["项目", "计量类别", "账面价值", "项目", "计量类别", "账面价值"]]
        titleLength = len(titles)
        rowLength = len(dc["index"]) + titleLength
        columnLength = len(dc["columns"])
        table = createBorderedTable(document, rowLength, columnLength)
        addCombineTableTitle(table, titles)
        addCombineTableContent(table, dc, titleLength)
    else:
        df = pd.read_excel(currentpath, sheet_name="首次执行日前后金融资产分类和计量对比表")
        dc = df.to_dict("split")
        titles = [["原金融工具准则", "nan", "nan", "新金融工具准则 ", "nan", "nan"], ["项目", "计量类别", "账面价值", "项目", "计量类别", "账面价值"]]
        titleLength = len(titles)
        rowLength = len(dc["index"]) + titleLength
        columnLength = len(dc["columns"])
        table = createBorderedTable(document, rowLength, columnLength)
        addCombineTableTitle(table, titles)
        addCombineTableContent(table, dc, titleLength)


# 添加新金融工具准则
def addNewFinancialInstrumentsChange(document, numTitle, currentpath,parentpath, context):
    reportType = context["report_params"]["type"]
    # 报告日期
    reportDate = context["report_params"]["reportDate"]
    # 获取报告起始日
    startYear = reportDate[:4]
    companyType = context["report_params"]["companyType"]
    # 新金融工具准则
    addParagraph(document, numTitle, "paragraph")
    addParagraph(document,
                 "在新金融工具准则下所有已确认金融资产，其后续均按摊余成本或公允价值计量。在新金融工具准则施行日，以本公司该日既有事实和情况为基础评估管理金融资产的业务模式、以金融资产初始确认时的事实和情况为基础评估该金融资产上的合同现金流量特征，将金融资产分为三类：按摊余成本计量、按公允价值计量且其变动计入其他综合收益及按公允价值计量且其变动计入当期损益。其中，对于按公允价值计量且其变动计入其他综合收益的权益工具投资，当该金融资产终止确认时，之前计入其他综合收益的累计利得或损失将从其他综合收益转入留存收益，不计入当期损益。",
                 "paragraph")
    addParagraph(document,
                 "在新金融工具准则下，本公司以预期信用损失为基础，对以摊余成本计量的金融资产、以公允价值计量且其变动计入其他综合收益的债务工具投资、租赁应收款、合同资产及财务担保合同计提减值准备并确认信用减值损失。",
                 "paragraph")
    addParagraph(document,
                 "本公司追溯应用新金融工具准则，但对于分类和计量（含减值）涉及前期比较财务报表数据与新金融工具准则不一致的，本公司选择不进行重述。因此，对于首次执行该准则的累积影响数，本公司调整{}年年初留存收益或其他综合收益以及财务报表其他相关项目金额，{}年度的财务报表未予重述。执行新金融工具准则的主要变化和影响如下：".format(
                     startYear, int(startYear) - 1), "paragraph")
    for content in context["standardChange"]["newFinancialInstrumentsChange"]:
        addParagraph(document, content, "paragraph")
    # 首次执行日前后金融资产分类和计量对比表
    addLandscapeContent(document, addContent1, currentpath,parentpath, context)

    # ②首次执行日，原金融资产账面价值调整为按照新金融工具准则的规定进行分类和计量的新金融资产账面价值的调节表
    if companyType == "国有企业":
        addParagraph(document, "（2）首次执行日，原金融资产账面价值调整为按照新金融工具准则的规定进行分类和计量的新金融资产账面价值的调节表", "paragraph")
    else:
        addParagraph(document, "②首次执行日，原金融资产账面价值调整为按照新金融工具准则的规定进行分类和计量的新金融资产账面价值的调节表", "paragraph")
    if reportType == "合并":
        # 添加合并报表
        addParagraph(document, "对本集团的影响：".format(startYear), "paragraph")
        df = pd.read_excel(currentpath, sheet_name="新旧金融工具调节表")
        dc = df.to_dict("split")
        addTable(document, dc, style=3)
        #     添加单体报表
        addParagraph(document, "对本公司的影响：".format(startYear),
                     "paragraph")
        df = pd.read_excel(parentpath, sheet_name="新旧金融工具调节表")
        dc = df.to_dict("split")
        addTable(document, dc, style=3)
    else:
        #     添加单体报表
        df = pd.read_excel(currentpath, sheet_name="新旧金融工具调节表")
        dc = df.to_dict("split")
        addTable(document, dc, style=3)

    # ③首次执行日，金融资产减值准备调节表
    if companyType == "国有企业":
        addParagraph(document, "（3）首次执行日，金融资产减值准备调节表", "paragraph")
    else:
        addParagraph(document, "③首次执行日，金融资产减值准备调节表", "paragraph")
    if reportType == "合并":
        # 添加合并影响
        addParagraph(document, "对本集团的影响：".format(startYear),
                     "paragraph")
        df = pd.read_excel(currentpath, sheet_name="金融资产减值准备调节表")
        dc = df.to_dict("split")
        addTable(document, dc, style=3)
        #     添加单体影响
        addParagraph(document, "对本公司的影响： ".format(startYear),
                     "paragraph")
        df = pd.read_excel(parentpath, sheet_name="金融资产减值准备调节表")
        dc = df.to_dict("split")
        addTable(document, dc, style=3)
    else:
        # 添加单体影响
        df = pd.read_excel(currentpath, sheet_name="金融资产减值准备调节表")
        dc = df.to_dict("split")
        addTable(document, dc, style=3)

    # 对2019年1月1日留存收益和其他综合收益的影响
    if companyType == "国有企业":
        addParagraph(document, "（4）对{}年1月1日留存收益和其他综合收益的影响".format(startYear), "paragraph")
    else:
        addParagraph(document, "④对{}年1月1日留存收益和其他综合收益的影响".format(startYear), "paragraph")
    if reportType == "合并":
        # 添加合并影响
        addParagraph(document, "对本集团的影响： ",
                     "paragraph")
        df = pd.read_excel(currentpath, sheet_name="新金融工具准则对期初留存收益和其他综合收益的影响")
        dc = df.to_dict("split")
        addTable(document, dc, style=3)
        #     添加单体影响
        addParagraph(document, "对本公司的影响： ",
                     "paragraph")
        df = pd.read_excel(parentpath, sheet_name="新金融工具准则对期初留存收益和其他综合收益的影响")
        dc = df.to_dict("split")
        addTable(document, dc, style=3)
    else:
        #  添加单体影响
        df = pd.read_excel(currentpath, sheet_name="新金融工具准则对期初留存收益和其他综合收益的影响")
        dc = df.to_dict("split")
        addTable(document, dc, style=3)


# 添加新收入准则
def addNewIncomeCriteria(document, numTitle, currentpath,parentpath, context):
    companyTye = context["report_params"]["companyType"]
    reportType = context["report_params"]["type"]
    # 报告日期
    reportDate = context["report_params"]["reportDate"]
    # 获取报告起始日
    startYear = reportDate[:4]
    # 新收入准则
    addParagraph(document, numTitle, "paragraph")
    addParagraph(document,
                 "新收入准则为规范与客户之间的合同产生的收入建立了新的收入确认模型。为执行新收入准则，本公司重新评估主要合同收入的确认和计量、核算和列报等方面。首次执行的累积影响金额调整首次执行当期期初（即{}年1月1日）的留存收益及财务报表其他相关项目金额，对可比期间信息不予调整。".format(
                     startYear), "paragraph")
    addParagraph(document, "执行新收入准则的主要变化和影响如下：", "paragraph")
    for content in context["standardChange"]["newIncomeCriteria"]:
        addParagraph(document, content, "paragraph")
    # （1）添加对期初报表的影响
    if companyTye == "国有企业":
        addParagraph(document, "（1）对{}年1月1日财务报表的影响：".format(startYear), "paragraph")
    else:
        addParagraph(document, "①对{}年1月1日财务报表的影响：".format(startYear), "paragraph")
    if reportType == "合并":
        # 添加合并报表
        # 添加合并影响
        addParagraph(document, "对本集团的影响： ",
                     "paragraph")
        df = pd.read_excel(currentpath, sheet_name="新收入准则对期初财务报表的影响")
        dc = df.to_dict("split")
        addTable(document, dc, style=3)
        #     添加单体影响
        addParagraph(document, "对本公司的影响： ",
                     "paragraph")
        df = pd.read_excel(parentpath, sheet_name="新收入准则对期初财务报表的影响")
        dc = df.to_dict("split")
        addTable(document, dc, style=3)
    else:
        # 添加单体报表
        df = pd.read_excel(parentpath, sheet_name="新收入准则对期初财务报表的影响")
        dc = df.to_dict("split")
        addTable(document, dc, style=3)
    # ②对2019年末资产负债表的影响
    if companyTye == "国有企业":
        addParagraph(document, "（2）对{}年末资产负债表的影响：".format(startYear), "paragraph")
    else:
        addParagraph(document, "②对{}年末资产负债表的影响：".format(startYear), "paragraph")
    if reportType == "合并":
        # 添加合并报表
        addParagraph(document, "对本集团的影响： ",
                     "paragraph")
        df = pd.read_excel(currentpath, sheet_name="新收入准则对期末资产负债表的影响")
        dc = df.to_dict("split")
        addTable(document, dc, style=3)
        #     添加单体影响
        addParagraph(document, "对本公司的影响： ",
                     "paragraph")
        df = pd.read_excel(parentpath, sheet_name="新收入准则对期末资产负债表的影响")
        dc = df.to_dict("split")
        addTable(document, dc, style=3)
    else:
        # 添加单体报表
        df = pd.read_excel(parentpath, sheet_name="新收入准则对期末资产负债表的影响")
        dc = df.to_dict("split")
        addTable(document, dc, style=3)
    # ③对2019年度利润表的影响
    if companyTye == "国有企业":
        addParagraph(document, "（3）对{}年末利润表的影响：".format(startYear), "paragraph")
    else:
        addParagraph(document, "③对{}年末利润表的影响：".format(startYear), "paragraph")
    if reportType == "合并":
        # 添加合并报表
        addParagraph(document, "对本集团的影响： ",
                     "paragraph")
        df = pd.read_excel(currentpath, sheet_name="新收入准则对利润表的影响")
        dc = df.to_dict("split")
        addTable(document, dc, style=3)
        #     添加单体影响
        addParagraph(document, "对本公司的影响： ",
                     "paragraph")
        df = pd.read_excel(parentpath, sheet_name="新收入准则对利润表的影响")
        dc = df.to_dict("split")
        addTable(document, dc, style=3)
    else:
        # 添加单体报表
        df = pd.read_excel(parentpath, sheet_name="新收入准则对利润表的影响")
        dc = df.to_dict("split")
        addTable(document, dc, style=3)


def addNewLeaseCriteria(document, numTitle, currentpath,parentpath, context):
    reportType = context["report_params"]["type"]
    # 报告日期
    reportDate = context["report_params"]["reportDate"]
    # 获取报告起始日
    startYear = reportDate[:4]
    # 新租赁准则
    addParagraph(document, numTitle, "paragraph")
    addParagraph(document,
                 "根据新租赁准则的规定，对于首次执行日前已存在的合同，本公司选择不重新评估其是否为租赁或者包含租赁。对作为承租人的租赁合同，本公司选择仅对{}年1月1日尚未完成的租赁合同的累计影响数进行调整。首次执行的累积影响金额调整首次执行当期期初的留存收益及财务报表其他相关项目金额，对可比期间信息未予重述。其中，对首次执行日的融资租赁，本公司作为承租人按照融资租入资产和应付融资租赁款的原账面价值，分别计量使用权资产和租赁负债；对首次执行日的经营租赁，作为承租人根据剩余租赁付款额按首次执行日的增量借款利率折现的现值计量租赁负债；原租赁准则下按照权责发生制计提的应付未付租金，纳入剩余租赁付款额中。".format(
                     startYear), "paragraph")
    addParagraph(document,
                 "本公司根据每项租赁选择按照下列两者之一计量使用权资产：A、假设自租赁期开始日即采用新租赁准则的账面价值（采用首次执行日的增量借款利率作为折现率）；B、与租赁负债相等的金额，并根据预付租金进行必要调整。并按照《企业会计准则第8号——资产减值》的规定，对使用权资产进行减值测试并进行相应会计处理。",
                 "paragraph")
    addParagraph(document, "执行新租赁准则的主要变化和影响如下：", "paragraph")
    for content in context["standardChange"]["newLeaseCriteria"]:
        addParagraph(document, content, "paragraph")
    # 上述会计政策变更对2019年1月1日合并财务报表的影响
    addParagraph(document, "上述会计政策变更对{}年1月1日财务报表的影响".format(startYear), "paragraph")
    if reportType == "合并":
        # 添加合并报表
        addParagraph(document, "对本集团的影响： ",
                     "paragraph")
        df = pd.read_excel(currentpath, sheet_name="新租赁准则对期初报表的影响")
        dc = df.to_dict("split")
        addTable(document, dc, style=3)
        #     添加单体影响
        addParagraph(document, "对本公司的影响： ",
                     "paragraph")
        df = pd.read_excel(parentpath, sheet_name="新租赁准则对期初报表的影响")
        dc = df.to_dict("split")
        addTable(document, dc, style=3)
    else:
        # 添加单体报表
        df = pd.read_excel(parentpath, sheet_name="新租赁准则对期初报表的影响")
        dc = df.to_dict("split")
        addTable(document, dc, style=3)
    addParagraph(document, "公司于{}年1月1日计入资产负债表的租赁负债所采用的增量借款利率的加权平均值为{}%。".format(startYear, context["standardChange"][
        "incrementalBorrowingRate"]), "paragraph")
    if reportType == "合并":
        addParagraph(document, "于{}年1月1日,本集团及本公司将原租赁准则下披露的尚未支付的最低经营租赁付款额调整为新租赁准则下确认的租赁负债的调节表如下：".format(startYear),
                     "paragraph")
        addParagraph(document, "对本集团的影响： ",
                     "paragraph")
        df = pd.read_excel(currentpath, sheet_name="最低经营租赁付款额与租赁负债调节表")
        dc = df.to_dict("split")
        addTable(document, dc, style=3)
        #     添加单体影响
        addParagraph(document, "对本公司的影响： ",
                     "paragraph")
        df = pd.read_excel(parentpath, sheet_name="最低经营租赁付款额与租赁负债调节表")
        dc = df.to_dict("split")
        addTable(document, dc, style=3)
    else:
        addParagraph(document, "于{}年1月1日,本公司将原租赁准则下披露的尚未支付的最低经营租赁付款额调整为新租赁准则下确认的租赁负债的调节表如下：".format(startYear),
                     "paragraph")
        df = pd.read_excel(parentpath, sheet_name="最低经营租赁付款额与租赁负债调节表")
        dc = df.to_dict("split")
        addTable(document, dc, style=3)


# 入口函数
def addChange(document, num, context, currentpath,parentpath):
    # 公司类型
    companyType = context["report_params"]["companyType"]
    # 报告日期
    reportDate = context["report_params"]["reportDate"]
    # 获取报告起始日
    startYear = reportDate[:4]
    # 合并报表还是单体报表
    reportType = context["report_params"]["type"]

    if companyType == "国有企业":
        addTitle(document, "五、会计政策、会计估计变更以及差错更正的说明", 1, False)
        addTitle(document, "（一）会计政策变更", 2, True)

        implementionOfNewStandards = {
            "新金融工具准则": "《企业会计准则第22号——金融工具确认和计量（2017年修订）》（财会〔2017〕7号）、《企业会计准则第23号——金融资产转移（2017年修订）》（财会〔2017〕8号）、《企业会计准则第24号——套期会计（2017年修订）》（财会〔2017〕9号），于2017年5月2日发布了《企业会计准则第37号——金融工具列报（2017年修订）》（财会〔2017〕14号）（以下合称“新金融工具准则”）",
            "新收入准则": "《企业会计准则第14号——收入（2017年修订）》（财会〔2017〕22号）（以下简称“新收入准则”）",
            "新租赁准则": "《企业会计准则第21号——租赁（2018年修订）》（财会[2018]35号）（以下简称“新租赁准则”）"}
        newStandardFuncs = {
            "新金融工具准则": addNewFinancialInstrumentsChange,
            "新收入准则": addNewIncomeCriteria,
            "新租赁准则": addNewLeaseCriteria,
        }

        if len(context["standardChange"]["implementationOfNewStandardsInThisPeriod"]) > 0:
            # 新准则
            desc = []
            for newStandard in context["standardChange"]["implementationOfNewStandardsInThisPeriod"]:
                desc.append(implementionOfNewStandards[newStandard])
            if len(desc) == 1:
                ImplementationOfNewStandardsDesc = "财政部于 2017年颁布了{}，本公司于{}年1月1日起开始执行前述准则。".format(desc[0], startYear)
            elif len(desc) == 2:
                ImplementationOfNewStandardsDesc = "财政部于 2017年颁布了{}，本公司于{}年1月1日起开始执行前述准则。".format(
                    desc[0] + "和" + desc[1],
                    startYear)
            elif len(desc) == 3:
                ImplementationOfNewStandardsDesc = "财政部于 2017年颁布了{}，本公司于{}年1月1日起开始执行前述准则。".format(
                    desc[0] + "," + desc[1] + "和" + desc[2], startYear)
            else:
                ImplementationOfNewStandardsDesc = ""

            addParagraph(document, ImplementationOfNewStandardsDesc, "paragraph")
            num = 1
            for newStandard in context["standardChange"]["implementationOfNewStandardsInThisPeriod"]:
                newStandardFuncs[newStandard](document, "{}、执行{}导致的会计政策变更".format(num, newStandard), currentpath,parentpath, context)

                num += 1
        else:
            addParagraph(document, "本公司{}年度无应披露的会计政策变更。".format(startYear),
                         "paragraph")

        addTitle(document, "（二）会计估计变更", 2, True)
        addParagraph(document, "本公司{}年度无应披露的会计估计变更。".format(startYear),
                     "paragraph")
        addTitle(document, "（三）重要前期差错更正", 2, True)
        addParagraph(document, "本公司{}年度无应披露的重要前期差错更正。".format(startYear), "paragraph")
    else:
        addTitle(document, "（{}）会计政策、会计估计变更以及差错更正的说明".format(to_chinese(num)), 2, True)
        addParagraph(document, "1、会计政策变更", "paragraph")

        implementionOfNewStandards = {
            "新金融工具准则": "《企业会计准则第22号——金融工具确认和计量（2017年修订）》（财会〔2017〕7号）、《企业会计准则第23号——金融资产转移（2017年修订）》（财会〔2017〕8号）、《企业会计准则第24号——套期会计（2017年修订）》（财会〔2017〕9号），于2017年5月2日发布了《企业会计准则第37号——金融工具列报（2017年修订）》（财会〔2017〕14号）（以下合称“新金融工具准则”）",
            "新收入准则": "《企业会计准则第14号——收入（2017年修订）》（财会〔2017〕22号）（以下简称“新收入准则”）",
            "新租赁准则": "《企业会计准则第21号——租赁（2018年修订）》（财会[2018]35号）（以下简称“新租赁准则”）"}
        newStandardFuncs = {
            "新金融工具准则": addNewFinancialInstrumentsChange,
            "新收入准则": addNewIncomeCriteria,
            "新租赁准则": addNewLeaseCriteria,
        }

        if len(context["standardChange"]["implementationOfNewStandardsInThisPeriod"]) > 0:
            # 新准则
            desc = []
            for newStandard in context["standardChange"]["implementationOfNewStandardsInThisPeriod"]:
                desc.append(implementionOfNewStandards[newStandard])
            if len(desc) == 1:
                ImplementationOfNewStandardsDesc = "财政部于 2017年颁布了{}，本公司于{}年1月1日起开始执行前述准则。".format(desc[0], startYear)
            elif len(desc) == 2:
                ImplementationOfNewStandardsDesc = "财政部于 2017年颁布了{}，本公司于{}年1月1日起开始执行前述准则。".format(
                    desc[0] + "和" + desc[1],
                    startYear)
            elif len(desc) == 3:
                ImplementationOfNewStandardsDesc = "财政部于 2017年颁布了{}，本公司于{}年1月1日起开始执行前述准则。".format(
                    desc[0] + "," + desc[1] + "和" + desc[2], startYear)
            else:
                ImplementationOfNewStandardsDesc = ""

            addParagraph(document, ImplementationOfNewStandardsDesc, "paragraph")
            num = 1
            for newStandard in context["standardChange"]["implementationOfNewStandardsInThisPeriod"]:
                newStandardFuncs[newStandard](document, "{}、执行{}导致的会计政策变更".format(num, newStandard), currentpath,
                                              parentpath, context)
                num += 1
        else:
            addParagraph(document, "本公司{}年度无应披露的会计政策变更。".format(startYear),
                         "paragraph")

        addParagraph(document, "2、会计估计变更", "paragraph")
        addParagraph(document, "本公司{}年度无应披露的会计估计变更。".format(startYear),
                     "paragraph")
        addParagraph(document, "3、重要前期差错更正", "paragraph")
        addParagraph(document, "本公司{}年度无应披露的重要前期差错更正。".format(startYear), "paragraph")


def test():
    from project.data import testcontext
    from project.constants import CURRENTPATH,PARENTPATH
    document = Document()
    # 设置中文标题
    setStyle(document)
    num = 5
    addChange(document, num, testcontext, CURRENTPATH,PARENTPATH)

    document.save("change.docx")


if __name__ == '__main__':
    test()
