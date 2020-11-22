# -*- coding: UTF-8 -*-

from docx import Document
import pandas as pd
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT


from project.data import report_params, standardChange,notes_params,noteAppend
from project.settings import setStyle
from project.utils import checkLeftSpace, addTitle, addCombineTableTitle,addContentToCombineTitle, addParagraph,createBorderedTable,addTable,setCell,addLandscapeContent


def excelTableToWord(sheetName,xlsxPath,style):
    df = pd.read_excel(xlsxPath, sheet_name=sheetName)
    dc = df.to_dict("split")
    addTable(document, dc, style=style)

# 报告日期
reportDate = report_params["reportDate"]
reportPeriod = report_params["reportPeriod"]
lastPeriod = str.replace(reportPeriod,reportPeriod[:4],str(int(reportPeriod[:4])-1))
# 人民币单位
currencyUnit = notes_params["currencyUnit"]
# 获取报告起始日
startYear = reportDate[:4]
# 合并报表还是单体报表
type = report_params["type"]
# 根据pandas读取的excle表格数据导入word
MODELPATH = "D:/auditReport/project/model.xlsx"
document = Document()
# 设置中文标题
setStyle(document)
addTitle(document, "八、合并财务报表重要项目的说明", 1, False)
basicDesc = "以下注释项目除特别注明之外，金额单位为{}；期初数指{}年12月31日财务报表数，期末数指{}财务报表数，本期指{}，上期指{}".format(currencyUnit,int(startYear)-1,reportDate,reportPeriod,lastPeriod)
addParagraph(document,basicDesc,"paragraph")
# （一）货币资金
addTitle(document, "（一）货币资金", 2, True)
excelTableToWord("货币资金",MODELPATH,style=2)

# 受限货币资金明细
df = pd.read_excel(MODELPATH, sheet_name="受限制的货币资金")
dc = df.to_dict("split")
if len(dc["data"])>1:
    addParagraph(document, "受限制的货币资金明细如下：", "paragraph")
    addTable(document, dc, style=2)

# （二）交易性金融资产
addTitle(document, "（二）交易性金融资产", 2, True)
excelTableToWord("交易性金融资产",MODELPATH,style=2)


# （三）以公允价值计量且其变动计入当期损益的金融资产
addTitle(document, "（三）以公允价值计量且其变动计入当期损益的金融资产", 2, True)
excelTableToWord("以公允价值计量且其变动计入当期损益的金融资产",MODELPATH,style=2)

# （四）衍生金融资产
addTitle(document, "（四）衍生金融资产", 2, True)
excelTableToWord("衍生金融资产",MODELPATH,style=2)


# （五）应收票据
# 为表格添加合并标题，最后一列向上合并
def addCombineTitleSpecialLast(titles,table):
    for i,row in enumerate(titles):
        for j,cellText in enumerate(row):
            cell = table.cell(i,j)
            if cellText=="nan" :
                if j>=1 and j!=len(row)-1:
                    if checkLeftSpace(row,j):
                        cell.merge(table.cell(i - 1, j))
                    else:
                        cell.merge(table.cell(i,j-1))
                else:
                    cell.merge(table.cell(i-1,j))
            else:
                setCell(cell,cellText,WD_PARAGRAPH_ALIGNMENT.CENTER)
addTitle(document, "（五）应收票据", 2, True)
# 老金融工具准则
addParagraph(document,"1、应收票据分类","paragraph")
if noteAppend["newFinancialInstruments"] == 0:
    excelTableToWord("应收票据分类原金融工具准则", MODELPATH, style=2)
else:
    df = pd.read_excel(MODELPATH, sheet_name="应收票据分类新金融工具准则")
    dc = df.to_dict("split")
    titles = [["种类", "期末数", "nan", "nan", "期初数", "nan", "nan"],
              ["nan", "账面余额", "坏账准备", "账面价值", "账面余额", "坏账准备", "账面价值"]]
    titleLength = len(titles)
    rowLength = len(dc["index"]) + titleLength
    columnLength = len(dc["columns"])
    table = createBorderedTable(document, rowLength, columnLength)
    addCombineTableTitle(table, titles)
    addContentToCombineTitle(document, dc, table, titleLength, style=2)

addParagraph(document,"2、期末已质押的应收票据","paragraph")
excelTableToWord("已质押应收票据", MODELPATH, style=2)

addParagraph(document,"3、期末已背书或贴现且在资产负债表日尚未到期的应收票据","paragraph")
excelTableToWord("已背书或贴现且在资产负债表日尚未到期的应收票据", MODELPATH, style=2)

addParagraph(document,"4、期末因出票人未履约而转为应收账款的票据","paragraph")
excelTableToWord("已背书或贴现且在资产负债表日尚未到期的应收票据", MODELPATH, style=2)

if noteAppend["newFinancialInstruments"] == 1:

    addParagraph(document,"5、期末单项计提坏账准备的应收票据","paragraph")
    excelTableToWord("期末单项计提坏账准备的应收票据新金融工具准则", MODELPATH, style=2)

    addParagraph(document, "6、采用组合计提坏账准备的应收票据", "paragraph")
    df = pd.read_excel(MODELPATH, sheet_name="采用组合计提坏账准备的应收票据新金融工具准则")
    dc = df.to_dict("split")
    titles = [["项目", "期末数", "nan", "nan"],
              ["nan", "账面余额", "坏账准备", "计提比例(%)"]]
    titleLength = len(titles)
    rowLength = len(dc["index"]) + titleLength
    columnLength = len(dc["columns"])
    table = createBorderedTable(document, rowLength, columnLength)
    addCombineTableTitle(table, titles)
    addContentToCombineTitle(document, dc, table, titleLength, style=2)

    addParagraph(document, "7、坏账准备变动明细情况", "paragraph")
    df = pd.read_excel(MODELPATH, sheet_name="应收票据坏账准备变动明细情况新金融工具准则")
    dc = df.to_dict("split")
    titles = [["项目", "期初数", "本期增加", "nan", "本期减少","nan","nan","期末数"],
              ["nan", "nan", "计提", "其他","转回","核销","其他","nan"]]
    titleLength = len(titles)
    rowLength = len(dc["index"]) + titleLength
    columnLength = len(dc["columns"])
    table = createBorderedTable(document, rowLength, columnLength)
    addCombineTitleSpecialLast(titles,table)
    addContentToCombineTitle(document, dc, table, titleLength, style=2)

    addParagraph(document, "8、本期重要的坏账准备收回或转回情况", "paragraph")
    excelTableToWord("本期重要的应收票据坏账准备收回或转回情况新金融工具准则", MODELPATH, style=2)

    addParagraph(document, "9、本期实际核销的应收票据情况", "paragraph")
    excelTableToWord("本期实际核销的应收票据情况新金融工具准则", MODELPATH, style=2)
# 应收账款
addTitle(document, "（六）应收账款", 2, True)
# 是否向上合并单元格
def isCombine(i,j,combineCells):
    for cell in combineCells:
        if i==cell[0] and j == cell[1]:
            return True
    return False

# 为表格添加合并标题，最后一列向上合并
def addCombineTitleSpecialReceivable(titles,table,combineCells):
    for i,row in enumerate(titles):
        for j,cellText in enumerate(row):
            cell = table.cell(i,j)
            if cellText=="nan" :
                # 最后一列，最后一行
                if isCombine(i,j,combineCells):
                    cell.merge(table.cell(i - 1, j))
                    continue
                if j>=1:
                    if checkLeftSpace(row,j):
                        cell.merge(table.cell(i - 1, j))
                    else:
                        cell.merge(table.cell(i,j-1))
                else:
                    cell.merge(table.cell(i-1,j))
            else:
                setCell(cell,cellText,WD_PARAGRAPH_ALIGNMENT.CENTER)
# 老准则
# 非首次执行新准则
# 本期首次执行新准则
if noteAppend["newFinancialInstruments"] == 0:
    # 原金融工具准则
    df = pd.read_excel(MODELPATH, sheet_name="应收账款期末数原金融工具准则")
    dc = df.to_dict("split")
    titles = [["种类", "期末数", "nan", "nan", "nan", "nan"],
              ["nan", "账面余额", "nan", "坏账准备", "nan", "账面价值"],
              ["nan", "金额", "比例(%)", "金额", "计提比例(%)", "nan"],
              ]
    titleLength = len(titles)
    rowLength = len(dc["index"]) + titleLength
    columnLength = len(dc["columns"])
    table = createBorderedTable(document, rowLength, columnLength)
    addCombineTitleSpecialReceivable(titles,table,[[2,5]])
    addContentToCombineTitle(document, dc, table, titleLength, style=2)
    addParagraph(document,"（续）","paragraph")
    df = pd.read_excel(MODELPATH, sheet_name="应收账款期初数新金融工具准则")
    dc = df.to_dict("split")
    rowLength = len(dc["index"]) + titleLength
    columnLength = len(dc["columns"])
    table = createBorderedTable(document, rowLength, columnLength)
    addCombineTitleSpecialReceivable(titles, table,[[2,5]])
    addContentToCombineTitle(document, dc, table, titleLength, style=2)

    addParagraph(document,"1、期末单项金额重大并单项计提坏账准备的应收账款","paragraph")
    excelTableToWord("期末单项计提坏账准备的应收账款", MODELPATH, style=2)

    addParagraph(document,"2、按信用风险特征组合计提坏账准备的应收账款","paragraph")
    addParagraph(document,"（1）采用账龄分析法计提坏账准备的应收账款","paragraph")
    df = pd.read_excel(MODELPATH, sheet_name="采用账龄分析法计提坏账准备的应收账款原准则")
    dc = df.to_dict("split")
    titles = [["账龄", "期末数", "nan", "nan","期初数", "nan", "nan"],
              ["nan", "账面余额", "nan", "坏账准备", "账面余额", "nan", "坏账准备"],
              ["nan", "金额", "比例(%)","nan", "金额", "比例(%)", "nan"],
              ]
    titleLength = len(titles)
    rowLength = len(dc["index"]) + titleLength
    columnLength = len(dc["columns"])
    table = createBorderedTable(document, rowLength, columnLength)
    addCombineTitleSpecialReceivable(titles, table,[[2,3],[2,6]])
    addContentToCombineTitle(document, dc, table, titleLength, style=2)

    addParagraph(document,"（2）采用其他组合方法计提坏账准备的应收账款","paragraph")
    df = pd.read_excel(MODELPATH, sheet_name="采用其他组合方法计提坏账准备的应收账款原准则")
    dc = df.to_dict("split")
    titles = [["组合名称", "期末数", "nan", "nan", "期初数", "nan", "nan"],
              ["nan", "账面余额", "计提比例（%）", "坏账准备", "账面余额", "计提比例（%）", "坏账准备"],
              ]
    titleLength = len(titles)
    rowLength = len(dc["index"]) + titleLength
    columnLength = len(dc["columns"])
    table = createBorderedTable(document, rowLength, columnLength)
    addCombineTitleSpecialReceivable(titles, table, [])
    addContentToCombineTitle(document, dc, table, titleLength, style=2)

    addParagraph(document,"3、期末单项金额虽不重大但单项计提坏账准备的应收账款","paragraph")
    excelTableToWord("期末单项金额虽不重大但单项计提坏账准备的应收账款原准则", MODELPATH, style=2)

    addParagraph(document,"4、本期重要的坏账准备收回或转回情况","paragraph")
    excelTableToWord("收回或转回的坏账准备情况", MODELPATH, style=2)

    addParagraph(document,"5、本期实际核销的应收账款情况","paragraph")
    excelTableToWord("本年实际核销的应收账款情况", MODELPATH, style=2)

    addParagraph(document,"6、按欠款方归集的年末余额前五名的应收账款情况","paragraph")
    excelTableToWord("按欠款方归集的年末余额前五名的应收账款情况", MODELPATH, style=2)

    addParagraph(document,"7、由金融资产转移而终止确认的应收账款","paragraph")
    excelTableToWord("由金融资产转移而终止确认的应收账款", MODELPATH, style=2)

    addParagraph(document,"8、转移应收账款且继续涉入形成的资产、负债","paragraph")
    excelTableToWord("转移应收账款且继续涉入形成的资产负债", MODELPATH, style=2)


else:
    if "新金融工具准则" in standardChange["implementationOfNewStandardsInThisPeriod"]:
        # 首次执行新金融工具准则
        df = pd.read_excel(MODELPATH, sheet_name="应收账款期末数首次新金融工具准则")
        dc = df.to_dict("split")
        titles = [["种类", "期末数", "nan", "nan", "nan", "nan"],
                  ["nan", "账面余额", "nan", "坏账准备", "nan", "账面价值"],
                  ["nan", "金额", "比例(%)", "金额", "计提比例(%)", "nan"],
                  ]
        titleLength = len(titles)
        rowLength = len(dc["index"]) + titleLength
        columnLength = len(dc["columns"])
        table = createBorderedTable(document, rowLength, columnLength)
        addCombineTitleSpecialReceivable(titles, table,[[2,5]])
        addContentToCombineTitle(document, dc, table, titleLength, style=2)
        addParagraph(document, "（续）", "paragraph")
        df = pd.read_excel(MODELPATH, sheet_name="应收账款期初数首次新金融工具准则")
        dc = df.to_dict("split")
        rowLength = len(dc["index"]) + titleLength
        columnLength = len(dc["columns"])
        table = createBorderedTable(document, rowLength, columnLength)
        addCombineTitleSpecialReceivable(titles, table,[[2,5]])
        addContentToCombineTitle(document, dc, table, titleLength, style=2)

        addParagraph(document, "1、期末单项计提坏账准备的应收账款", "paragraph")
        excelTableToWord("期末单项计提坏账准备的应收账款", MODELPATH, style=2)

        addParagraph(document, "2、采用组合计提坏账准备的应收账款", "paragraph")
        df = pd.read_excel(MODELPATH, sheet_name="采用组合计提坏账准备的应收账款首次执行")
        dc = df.to_dict("split")
        titles = [["组合名称", "期末数", "nan", "nan"],
                  ["nan", "账面余额", "坏账准备", "计提比例(%)"],
                  ]
        titleLength = len(titles)
        rowLength = len(dc["index"]) + titleLength
        columnLength = len(dc["columns"])
        table = createBorderedTable(document, rowLength, columnLength)
        addCombineTitleSpecialReceivable(titles, table, [])
        addContentToCombineTitle(document, dc, table, titleLength, style=2)

        for row in dc["data"][:-1]:
            combinationName = row[0] + "首次执行"
            addParagraph(document, "{}:".format(row[0]), "paragraph")
            df = pd.read_excel(MODELPATH, sheet_name=combinationName)
            dc = df.to_dict("split")
            titles = [["项目", "期末数", "nan", "nan"],
                      ["nan", "账面余额", "坏账准备", "计提比例(%)"],
                      ]
            titleLength = len(titles)
            rowLength = len(dc["index"]) + titleLength
            columnLength = len(dc["columns"])
            table = createBorderedTable(document, rowLength, columnLength)
            addCombineTitleSpecialReceivable(titles, table, [])
            addContentToCombineTitle(document, dc, table, titleLength, style=2)

        addParagraph(document, "3、坏账准备变动明细情况", "paragraph")
        df = pd.read_excel(MODELPATH, sheet_name="应收账款坏账准备变动明细情况新金融工具准则 ")
        dc = df.to_dict("split")
        titles = [["项目", "期初数", "本期增加", "nan", "本期减少", "nan", "nan", "期末数"],
                  ["nan", "nan", "计提", "其他", "转回", "核销", "其他", "nan"]]
        titleLength = len(titles)
        rowLength = len(dc["index"]) + titleLength
        columnLength = len(dc["columns"])
        table = createBorderedTable(document, rowLength, columnLength)
        addCombineTitleSpecialLast(titles, table)
        addContentToCombineTitle(document, dc, table, titleLength, style=2)

        addParagraph(document, "4、本期重要的坏账准备收回或转回情况", "paragraph")
        excelTableToWord("收回或转回的坏账准备情况", MODELPATH, style=2)

        addParagraph(document, "5、本期实际核销的应收账款情况", "paragraph")
        excelTableToWord("本年实际核销的应收账款情况", MODELPATH, style=2)

        addParagraph(document, "6、按欠款方归集的年末余额前五名的应收账款情况", "paragraph")
        excelTableToWord("按欠款方归集的年末余额前五名的应收账款情况", MODELPATH, style=2)

        addParagraph(document, "7、由金融资产转移而终止确认的应收账款", "paragraph")
        excelTableToWord("由金融资产转移而终止确认的应收账款", MODELPATH, style=2)

        addParagraph(document, "8、转移应收账款且继续涉入形成的资产、负债", "paragraph")
        excelTableToWord("转移应收账款且继续涉入形成的资产负债", MODELPATH, style=2)

    else:
        # 新金融工具准则
        df = pd.read_excel(MODELPATH, sheet_name="应收账款期末数新金融工具准则")
        dc = df.to_dict("split")
        titles = [["种类", "期末数", "nan", "nan", "nan", "nan"],
                  ["nan", "账面余额", "nan", "坏账准备", "nan", "账面价值"],
                  ["nan", "金额", "比例(%)", "金额", "计提比例(%)", "nan"],
                  ]
        titleLength = len(titles)
        rowLength = len(dc["index"]) + titleLength
        columnLength = len(dc["columns"])
        table = createBorderedTable(document, rowLength, columnLength)
        addCombineTitleSpecialReceivable(titles, table,[[2,5]])
        addContentToCombineTitle(document, dc, table, titleLength, style=2)
        addParagraph(document, "（续）", "paragraph")
        df = pd.read_excel(MODELPATH, sheet_name="应收账款期初数新金融工具准则")
        dc = df.to_dict("split")
        rowLength = len(dc["index"]) + titleLength
        columnLength = len(dc["columns"])
        table = createBorderedTable(document, rowLength, columnLength)
        addCombineTitleSpecialReceivable(titles, table,[[2,5]])
        addContentToCombineTitle(document, dc, table, titleLength, style=2)

        addParagraph(document, "1、期末单项计提坏账准备的应收账款", "paragraph")
        excelTableToWord("期末单项计提坏账准备的应收账款", MODELPATH, style=2)

        addParagraph(document, "2、采用组合计提坏账准备的应收账款", "paragraph")
        df = pd.read_excel(MODELPATH, sheet_name="采用组合计提坏账准备的应收账款")
        dc = df.to_dict("split")
        titles = [["组合名称", "期末数", "nan", "nan","期初数","nan","nan"],
                  ["nan", "账面余额", "坏账准备", "计提比例(%)","账面余额", "坏账准备", "计提比例(%)"],
                  ]
        titleLength = len(titles)
        rowLength = len(dc["index"]) + titleLength
        columnLength = len(dc["columns"])
        table = createBorderedTable(document, rowLength, columnLength)
        addCombineTitleSpecialReceivable(titles, table, [])
        addContentToCombineTitle(document, dc, table, titleLength, style=2)

        for row in dc["data"][:-1]:
            combinationName = row[0]
            addParagraph(document, "{}:".format(row[0]), "paragraph")
            df = pd.read_excel(MODELPATH, sheet_name=combinationName)
            dc = df.to_dict("split")
            titles = [["项目", "期末数", "nan", "nan", "期初数", "nan", "nan"],
                      ["nan", "账面余额", "坏账准备", "计提比例(%)", "账面余额", "坏账准备", "计提比例(%)"],
                      ]
            titleLength = len(titles)
            rowLength = len(dc["index"]) + titleLength
            columnLength = len(dc["columns"])
            table = createBorderedTable(document, rowLength, columnLength)
            addCombineTitleSpecialReceivable(titles, table, [])
            addContentToCombineTitle(document, dc, table, titleLength, style=2)

    addParagraph(document, "3、坏账准备变动明细情况", "paragraph")
    df = pd.read_excel(MODELPATH, sheet_name="应收账款坏账准备变动明细情况新金融工具准则 ")
    dc = df.to_dict("split")
    titles = [["项目", "期初数", "本期增加", "nan", "本期减少", "nan", "nan", "期末数"],
              ["nan", "nan", "计提", "其他", "转回", "核销", "其他", "nan"]]
    titleLength = len(titles)
    rowLength = len(dc["index"]) + titleLength
    columnLength = len(dc["columns"])
    table = createBorderedTable(document, rowLength, columnLength)
    addCombineTitleSpecialLast(titles, table)
    addContentToCombineTitle(document, dc, table, titleLength, style=2)

    addParagraph(document, "4、本期重要的坏账准备收回或转回情况", "paragraph")
    excelTableToWord("收回或转回的坏账准备情况", MODELPATH, style=2)

    addParagraph(document, "5、本期实际核销的应收账款情况", "paragraph")
    excelTableToWord("本年实际核销的应收账款情况", MODELPATH, style=2)

    addParagraph(document, "6、按欠款方归集的年末余额前五名的应收账款情况", "paragraph")
    excelTableToWord("按欠款方归集的年末余额前五名的应收账款情况", MODELPATH, style=2)

    addParagraph(document, "7、由金融资产转移而终止确认的应收账款", "paragraph")
    excelTableToWord("由金融资产转移而终止确认的应收账款", MODELPATH, style=2)

    addParagraph(document, "8、转移应收账款且继续涉入形成的资产、负债", "paragraph")
    excelTableToWord("转移应收账款且继续涉入形成的资产负债", MODELPATH, style=2)

addTitle(document, "（七）应收款项融资", 2, True)
excelTableToWord("应收款项融资", MODELPATH, style=2)
addParagraph(document,"本公司的应收款项融资主要为根据日常资金管理需要，预计通过转让、贴现或背书回款并终止确认的应收账款和银行承兑汇票。","paragraph")
addParagraph(document,"本公司无单项计提减值准备的银行承兑汇票。于{}，本公司按照整个存续期预期信用损失计量坏账准备，本公司认为所持有的银行承兑汇票不存在重大信用风险，不会因银行违约而产生重大损失。".format(reportDate),"paragraph")
addParagraph(document,"于{}，本公司列示于应收款项融资已转让、已背书或已贴现但尚未到期的应收票据和应收账款如下：".format(reportDate),"paragraph")
excelTableToWord("应收款项融资已转让已背书或已贴现未到期", MODELPATH, style=2)

addTitle(document, "（八）预付款项", 2, True)
addParagraph(document, "1、	按账龄列示", "paragraph")
df = pd.read_excel(MODELPATH, sheet_name="预付账款账龄明细")
dc = df.to_dict("split")
titles = [["账  龄", "期末数", "nan", "nan", "nan", "期初数", "nan", "nan", "nan"],
          ["nan", "账面余额", "比例(%)", "减值准备", "账面价值", "账面余额", "比例(%)", "减值准备", "账面价值"]]
titleLength = len(titles)
rowLength = len(dc["index"]) + titleLength
columnLength = len(dc["columns"])
table = createBorderedTable(document, rowLength, columnLength)
addCombineTitleSpecialLast(titles, table)
addContentToCombineTitle(document, dc, table, titleLength, style=2)
addParagraph(document, "2、	账龄超过1年的大额预付款项情况", "paragraph")
excelTableToWord("账龄超过1年的大额预付款项情况", MODELPATH, style=2)
addParagraph(document, "3、按欠款方归集的年末余额前五名的预付账款情况", "paragraph")
excelTableToWord("按欠款方归集的年末余额前五名的预付账款情况", MODELPATH, style=2)


addTitle(document, "（九）其他应收款", 2, True)
if noteAppend["newFinancialInstruments"] == 0:
    # 原金融工具准则
    excelTableToWord("其他应收款原准则", MODELPATH, style=2)

    addParagraph(document,"1、应收利息","paragraph")
    addParagraph(document,"（1）应收利息分类","paragraph")
    excelTableToWord("应收利息分类", MODELPATH, style=2)
    addParagraph(document,"（2）重要逾期利息","paragraph")
    excelTableToWord("重要逾期利息", MODELPATH, style=2)

    addParagraph(document,"2、应收股利","paragraph")
    excelTableToWord("应收股利明细", MODELPATH, style=2)

    addParagraph(document,"3、其他应收款项","paragraph")
    df = pd.read_excel(MODELPATH, sheet_name="其他应收款项期末明细原准则")
    dc = df.to_dict("split")
    titles = [["种类", "期末数", "nan", "nan", "nan", "nan"],
              ["nan", "账面余额", "nan", "坏账准备", "nan", "账面价值"],
              ["nan", "金额", "比例(%)", "金额", "计提比例(%)", "nan"],
              ]
    titleLength = len(titles)
    rowLength = len(dc["index"]) + titleLength
    columnLength = len(dc["columns"])
    table = createBorderedTable(document, rowLength, columnLength)
    addCombineTitleSpecialReceivable(titles,table,[[2,5]])
    addContentToCombineTitle(document, dc, table, titleLength, style=2)
    addParagraph(document,"（续）","paragraph")
    df = pd.read_excel(MODELPATH, sheet_name="其他应收款项期初明细原准则")
    dc = df.to_dict("split")
    rowLength = len(dc["index"]) + titleLength
    columnLength = len(dc["columns"])
    table = createBorderedTable(document, rowLength, columnLength)
    addCombineTitleSpecialReceivable(titles, table,[[2,5]])
    addContentToCombineTitle(document, dc, table, titleLength, style=2)

    addParagraph(document,"（1）年末单项金额重大并单项计提坏账准备的其他应收款项","paragraph")
    excelTableToWord("期末单项计提坏账准备的其他应收款", MODELPATH, style=2)

    addParagraph(document,"（2）按信用风险特征组合计提坏账准备的其他应收款项","paragraph")
    addParagraph(document,"①采用账龄分析法计提坏账准备的其他应收款项","paragraph")
    df = pd.read_excel(MODELPATH, sheet_name="采用账龄分析法计提坏账准备的其他应收款项原准则")
    dc = df.to_dict("split")
    titles = [["账龄", "期末数", "nan", "nan","期初数", "nan", "nan"],
              ["nan", "账面余额", "nan", "坏账准备", "账面余额", "nan", "坏账准备"],
              ["nan", "金额", "比例(%)","nan", "金额", "比例(%)", "nan"],
              ]
    titleLength = len(titles)
    rowLength = len(dc["index"]) + titleLength
    columnLength = len(dc["columns"])
    table = createBorderedTable(document, rowLength, columnLength)
    addCombineTitleSpecialReceivable(titles, table,[[2,3],[2,6]])
    addContentToCombineTitle(document, dc, table, titleLength, style=2)

    addParagraph(document,"②采用余额百分比法或其他组合方法计提坏账准备的其他应收款项","paragraph")
    df = pd.read_excel(MODELPATH, sheet_name="采用其他组合方法计提坏账准备的其他应收款原准则")
    dc = df.to_dict("split")
    titles = [["组合名称", "期末数", "nan", "nan", "期初数", "nan", "nan"],
              ["nan", "账面余额", "计提比例（%）", "坏账准备", "账面余额", "计提比例（%）", "坏账准备"],
              ]
    titleLength = len(titles)
    rowLength = len(dc["index"]) + titleLength
    columnLength = len(dc["columns"])
    table = createBorderedTable(document, rowLength, columnLength)
    addCombineTitleSpecialReceivable(titles, table, [])
    addContentToCombineTitle(document, dc, table, titleLength, style=2)

    addParagraph(document,"（3）期末单项金额虽不重大但单项计提坏账准备的其他应收款项","paragraph")
    excelTableToWord("期末单项金额虽不重大但单项计提坏账准备的其他应收款原准则", MODELPATH, style=2)

    addParagraph(document,"（4）本期重要的坏账准备收回或转回情况","paragraph")
    excelTableToWord("其他应收款收回或转回的坏账准备情况", MODELPATH, style=2)

    addParagraph(document,"（5）本期实际核销的其他应收款情况","paragraph")
    excelTableToWord("本年实际核销的其他应收款情况", MODELPATH, style=2)

    addParagraph(document,"（6）按欠款方归集的年末金额前五名的其他应收款项情况","paragraph")
    excelTableToWord("按欠款方归集的年末金额前五名的其他应收款项情况", MODELPATH, style=2)

    addParagraph(document,"（7）由金融资产转移而终止确认的其他应收款项","paragraph")
    excelTableToWord("由金融资产转移而终止确认的其他应收款项", MODELPATH, style=2)

    addParagraph(document,"（8）转移其他应收款且继续涉入形成的资产、负债","paragraph")
    excelTableToWord("转移其他应收款且继续涉入形成的资产负债", MODELPATH, style=2)

    addParagraph(document, "（9）按应收金额确认的政府补助", "paragraph")
    excelTableToWord("按应收金额确认的政府补助", MODELPATH, style=2)


else:
    if "新金融工具准则" in standardChange["implementationOfNewStandardsInThisPeriod"]:
        # 首次执行新金融工具准则
        addParagraph(document,"1、明细情况","paragraph")
        addParagraph(document,"（1）类别明细情况","paragraph")
        df = pd.read_excel(MODELPATH, sheet_name="其他应收款期末数首次新金融工具准则")
        dc = df.to_dict("split")
        titles = [["种类", "期末数", "nan", "nan", "nan", "nan"],
                  ["nan", "账面余额", "nan", "坏账准备", "nan", "账面价值"],
                  ["nan", "金额", "比例(%)", "金额", "计提比例(%)", "nan"],
                  ]
        titleLength = len(titles)
        rowLength = len(dc["index"]) + titleLength
        columnLength = len(dc["columns"])
        table = createBorderedTable(document, rowLength, columnLength)
        addCombineTitleSpecialReceivable(titles, table,[[2,5]])
        addContentToCombineTitle(document, dc, table, titleLength, style=2)
        addParagraph(document, "（续）", "paragraph")
        df = pd.read_excel(MODELPATH, sheet_name="其他应收款期初数首次新金融工具准则")
        dc = df.to_dict("split")
        rowLength = len(dc["index"]) + titleLength
        columnLength = len(dc["columns"])
        table = createBorderedTable(document, rowLength, columnLength)
        addCombineTitleSpecialReceivable(titles, table,[[2,5]])
        addContentToCombineTitle(document, dc, table, titleLength, style=2)

        addParagraph(document, "（2）期末单项计提坏账准备的其他应收款", "paragraph")
        excelTableToWord("期末单项计提坏账准备的其他应收款", MODELPATH, style=2)

        addParagraph(document, "（3）采用组合计提坏账准备的其他应收款", "paragraph")
        df = pd.read_excel(MODELPATH, sheet_name="采用组合计提坏账准备的其他应收款新金融工具准则")
        dc = df.to_dict("split")
        titles = [["组合名称", "期末数", "nan", "nan"],
                  ["nan", "账面余额", "坏账准备", "计提比例(%)"],
                  ]
        titleLength = len(titles)
        rowLength = len(dc["index"]) + titleLength
        columnLength = len(dc["columns"])
        table = createBorderedTable(document, rowLength, columnLength)
        addCombineTitleSpecialReceivable(titles, table, [])
        addContentToCombineTitle(document, dc, table, titleLength, style=2)

        addParagraph(document, "2、账龄情况", "paragraph")
        excelTableToWord("其他应收款账龄情况新金融工具准则", MODELPATH, style=2)

        addParagraph(document, "3、坏账准备变动明细情况", "paragraph")
        df = pd.read_excel(MODELPATH, sheet_name="其他应收款坏账准备变动情况新金融工具准则")
        dc = df.to_dict("split")
        titles = [["项目", "第一阶段", "第二阶段", "第三阶段", "合  计"],
                  ["nan", "未来12个月预期信用损失", "整个存续期预期信用损失(未发生信用减值)", "整个存续期预期信用损失(已发生信用减值)","nan"]]
        titleLength = len(titles)
        rowLength = len(dc["index"]) + titleLength
        columnLength = len(dc["columns"])
        table = createBorderedTable(document, rowLength, columnLength)
        addCombineTitleSpecialLast(titles, table)
        addContentToCombineTitle(document, dc, table, titleLength, style=2)

        addParagraph(document, "4、本期重要的坏账准备收回或转回情况", "paragraph")
        excelTableToWord("收回或转回的坏账准备情况", MODELPATH, style=2)

        addParagraph(document, "5、本期实际核销的应收账款情况", "paragraph")
        excelTableToWord("本年实际核销的应收账款情况", MODELPATH, style=2)

        addParagraph(document, "6、其他应收款款项性质分类情况", "paragraph")
        excelTableToWord("其他应收款按性质分类情况", MODELPATH, style=2)

        addParagraph(document, "7、重要逾期利息", "paragraph")
        excelTableToWord("重要逾期利息", MODELPATH, style=2)

        addParagraph(document, "8、应收股利明细情况", "paragraph")
        excelTableToWord("应收股利明细", MODELPATH, style=2)

        addParagraph(document, "9、按欠款方归集的年末余额前五名的其他应收款情况", "paragraph")
        excelTableToWord("按欠款方归集的年末金额前五名的其他应收款项情况", MODELPATH, style=2)

        addParagraph(document, "10、按应收金额确认的政府补助", "paragraph")
        excelTableToWord("按应收金额确认的政府补助", MODELPATH, style=2)

        addParagraph(document, "11、由金融资产转移而终止确认的应收账款", "paragraph")
        excelTableToWord("由金融资产转移而终止确认的应收账款", MODELPATH, style=2)

        addParagraph(document, "12、转移应收账款且继续涉入形成的资产、负债", "paragraph")
        excelTableToWord("转移应收账款且继续涉入形成的资产负债", MODELPATH, style=2)

    else:
        # 新金融工具准则
        addParagraph(document, "1、明细情况", "paragraph")
        addParagraph(document, "（1）类别明细情况", "paragraph")
        df = pd.read_excel(MODELPATH, sheet_name="其他应收款期末数新金融工具准则")
        dc = df.to_dict("split")
        titles = [["种类", "期末数", "nan", "nan", "nan", "nan"],
                  ["nan", "账面余额", "nan", "坏账准备", "nan", "账面价值"],
                  ["nan", "金额", "比例(%)", "金额", "计提比例(%)", "nan"],
                  ]
        titleLength = len(titles)
        rowLength = len(dc["index"]) + titleLength
        columnLength = len(dc["columns"])
        table = createBorderedTable(document, rowLength, columnLength)
        addCombineTitleSpecialReceivable(titles, table, [[2, 5]])
        addContentToCombineTitle(document, dc, table, titleLength, style=2)
        addParagraph(document, "（续）", "paragraph")
        df = pd.read_excel(MODELPATH, sheet_name="其他应收款期初数新金融工具准则")
        dc = df.to_dict("split")
        rowLength = len(dc["index"]) + titleLength
        columnLength = len(dc["columns"])
        table = createBorderedTable(document, rowLength, columnLength)
        addCombineTitleSpecialReceivable(titles, table, [[2, 5]])
        addContentToCombineTitle(document, dc, table, titleLength, style=2)

        addParagraph(document, "（2）期末单项计提坏账准备的其他应收款", "paragraph")
        excelTableToWord("期末单项计提坏账准备的其他应收款", MODELPATH, style=2)

        addParagraph(document, "（3）采用组合计提坏账准备的其他应收款", "paragraph")
        df = pd.read_excel(MODELPATH, sheet_name="采用组合计提坏账准备的其他应收款新金融工具准则")
        dc = df.to_dict("split")
        titles = [["组合名称", "期末数", "nan", "nan"],
                  ["nan", "账面余额", "坏账准备", "计提比例(%)"],
                  ]
        titleLength = len(titles)
        rowLength = len(dc["index"]) + titleLength
        columnLength = len(dc["columns"])
        table = createBorderedTable(document, rowLength, columnLength)
        addCombineTitleSpecialReceivable(titles, table, [])
        addContentToCombineTitle(document, dc, table, titleLength, style=2)

        addParagraph(document, "2、账龄情况", "paragraph")
        excelTableToWord("其他应收款账龄情况新金融工具准则", MODELPATH, style=2)

        addParagraph(document, "3、坏账准备变动明细情况", "paragraph")
        df = pd.read_excel(MODELPATH, sheet_name="其他应收款坏账准备变动情况新金融工具准则")
        dc = df.to_dict("split")
        titles = [["项目", "第一阶段", "第二阶段", "第三阶段", "合  计"],
                  ["nan", "未来12个月预期信用损失", "整个存续期预期信用损失(未发生信用减值)", "整个存续期预期信用损失(已发生信用减值)","nan"]]
        titleLength = len(titles)
        rowLength = len(dc["index"]) + titleLength
        columnLength = len(dc["columns"])
        table = createBorderedTable(document, rowLength, columnLength)
        addCombineTitleSpecialLast(titles, table)
        addContentToCombineTitle(document, dc, table, titleLength, style=2)

        addParagraph(document, "4、本期重要的坏账准备收回或转回情况", "paragraph")
        excelTableToWord("收回或转回的坏账准备情况", MODELPATH, style=2)

        addParagraph(document, "5、本期实际核销的应收账款情况", "paragraph")
        excelTableToWord("本年实际核销的应收账款情况", MODELPATH, style=2)

        addParagraph(document, "6、其他应收款款项性质分类情况", "paragraph")
        excelTableToWord("其他应收款按性质分类情况", MODELPATH, style=2)

        addParagraph(document, "7、重要逾期利息", "paragraph")
        excelTableToWord("重要逾期利息", MODELPATH, style=2)

        addParagraph(document, "8、应收股利明细情况", "paragraph")
        excelTableToWord("应收股利明细", MODELPATH, style=2)

        addParagraph(document, "9、按欠款方归集的年末余额前五名的其他应收款情况", "paragraph")
        excelTableToWord("按欠款方归集的年末金额前五名的其他应收款项情况", MODELPATH, style=2)

        addParagraph(document, "10、按应收金额确认的政府补助", "paragraph")
        excelTableToWord("按应收金额确认的政府补助", MODELPATH, style=2)

        addParagraph(document, "11、由金融资产转移而终止确认的应收账款", "paragraph")
        excelTableToWord("由金融资产转移而终止确认的应收账款", MODELPATH, style=2)

        addParagraph(document, "12、转移应收账款且继续涉入形成的资产、负债", "paragraph")
        excelTableToWord("转移应收账款且继续涉入形成的资产负债", MODELPATH, style=2)

addTitle(document, "（十）存货", 2, True)
addParagraph(document, "1、明细情况", "paragraph")
df = pd.read_excel(MODELPATH, sheet_name="存货明细情况")
dc = df.to_dict("split")
titles = [["项  目", "期末数", "nan", "nan", "期初数", "nan", "nan"],
          ["nan", "账面余额", "跌价准备", "账面价值","账面余额", "跌价准备", "账面价值"]]
titleLength = len(titles)
rowLength = len(dc["index"]) + titleLength
columnLength = len(dc["columns"])
table = createBorderedTable(document, rowLength, columnLength)
addCombineTableTitle(table,titles)
addContentToCombineTitle(document, dc, table, titleLength, style=2)

addParagraph(document,"房地产开发成本明细如下：","paragraph")
excelTableToWord("房地产开发成本", MODELPATH, style=2)

addParagraph(document,"房地产开发产品明细如下：","paragraph")
excelTableToWord("房地产开发产品", MODELPATH, style=2)

addParagraph(document,"合同履约成本明细如下：","paragraph")
excelTableToWord("合同履约成本", MODELPATH, style=2)

addParagraph(document, "2、存货跌价准备", "paragraph")
df = pd.read_excel(MODELPATH, sheet_name="存货跌价准备明细情况")
dc = df.to_dict("split")
titles = [["项  目", "期初数", "本期增加", "nan", "本期减少", "nan", "期末数"],
          ["nan", "nan", "计提", "其他","转回或转销", "其他", "nan"]]
titleLength = len(titles)
rowLength = len(dc["index"]) + titleLength
columnLength = len(dc["columns"])
table = createBorderedTable(document, rowLength, columnLength)
addCombineTitleSpecialReceivable(titles, table,[[1,6]])
addContentToCombineTitle(document, dc, table, titleLength, style=2)

addParagraph(document, "3、确定可变现净值的具体依据、本期转回或转销存货跌价准备的原因", "paragraph")
excelTableToWord("确定可变现净值的具体依据", MODELPATH, style=2)

addParagraph(document, "4、存货期末余额中借款费用资本化情况", "paragraph")
excelTableToWord("存货期末余额中借款费用资本化情况", MODELPATH, style=2)

addTitle(document, "（十一）合同资产", 2, True)
addParagraph(document,"1、合同资产情况","paragraph")
df = pd.read_excel(MODELPATH, sheet_name="存货跌价准备明细情况")
dc = df.to_dict("split")
titles = [["项  目", "期末数", "nan", "nan", "期初数", "nan", "nan"],
          ["nan", "账面余额", "减值准备", "账面价值","账面余额", "减值准备", "账面价值"]]
titleLength = len(titles)
rowLength = len(dc["index"]) + titleLength
columnLength = len(dc["columns"])
table = createBorderedTable(document, rowLength, columnLength)
addCombineTableTitle(table,titles)
addContentToCombineTitle(document, dc, table, titleLength, style=2)
addParagraph(document,"2、合同资产本期的重大变动","paragraph")
excelTableToWord("合同资产本期的重大变动", MODELPATH, style=2)
addParagraph(document,"3、合同资产减值准备","paragraph")
addParagraph(document,"（1）期末单项计提坏账准备的合同资产","paragraph")
excelTableToWord("期末单项计提坏账准备的合同资产", MODELPATH, style=2)
addParagraph(document,"（2）采用组合计提坏账准备的合同资产","paragraph")
excelTableToWord("采用组合计提坏账准备的合同资产", MODELPATH, style=2)

addTitle(document, "（十二）持有待售资产", 2, True)
addParagraph(document,"1、持有待售资产的基本情况","paragraph")
excelTableToWord("持有待售资产的基本情况", MODELPATH, style=2)
addParagraph(document,"2、持有待售资产减值准备情况","paragraph")
excelTableToWord("持有待售资产减值准备情况", MODELPATH, style=2)
addParagraph(document,"3、与持有待售的非流动资产或处置组有关的其他综合收益累计金额","paragraph")
excelTableToWord("与持有待售的非流动资产或处置组有关的其他综合收益累计金额", MODELPATH, style=2)
addParagraph(document,"4、本期不再划分为持有待售类别或从持有待售处置组中移除的情况","paragraph")
excelTableToWord("本期不再划分为持有待售类别或从持有待售处置组中移除的情况", MODELPATH, style=2)

addTitle(document, "（十三）一年内到期的非流动资产", 2, True)
excelTableToWord("一年内到期的非流动资产",MODELPATH,style=2)

addTitle(document, "（十四）其他流动资产", 2, True)
excelTableToWord("其他流动资产",MODELPATH,style=2)
addParagraph(document,"合同取得成本本期变动：","paragraph")
excelTableToWord("合同取得成本",MODELPATH,style=2)

addTitle(document, "（十五）债权投资", 2, True)
addParagraph(document,"1、明细情况","paragraph")
df = pd.read_excel(MODELPATH, sheet_name="债权投资期末数")
dc = df.to_dict("split")
titles = [["项  目", "期末数", "nan", "nan", "nan", "nan"],
          ["nan", "初始成本", "利息调整", "应计利息","减值准备","账面价值"]]
titleLength = len(titles)
rowLength = len(dc["index"]) + titleLength
columnLength = len(dc["columns"])
table = createBorderedTable(document, rowLength, columnLength)
addCombineTableTitle(table,titles)
addContentToCombineTitle(document, dc, table, titleLength, style=2)
addParagraph(document,"（续）","paragraph")
df = pd.read_excel(MODELPATH, sheet_name="债权投资期初数")
dc = df.to_dict("split")
titles = [["项  目", "期末数", "nan", "nan", "nan", "nan"],
          ["nan", "初始成本", "利息调整", "应计利息","减值准备","账面价值"]]
titleLength = len(titles)
rowLength = len(dc["index"]) + titleLength
columnLength = len(dc["columns"])
table = createBorderedTable(document, rowLength, columnLength)
addCombineTableTitle(table,titles)
addContentToCombineTitle(document, dc, table, titleLength, style=2)
addParagraph(document,"2、债权投资减值准备","paragraph")
df = pd.read_excel(MODELPATH, sheet_name="债权投资减值准备")
dc = df.to_dict("split")
titles = [["项目", "第一阶段", "第二阶段", "第三阶段", "合  计"],
          ["nan", "未来12个月预期信用损失", "整个存续期预期信用损失(未发生信用减值)", "整个存续期预期信用损失(已发生信用减值)","nan"]]
titleLength = len(titles)
rowLength = len(dc["index"]) + titleLength
columnLength = len(dc["columns"])
table = createBorderedTable(document, rowLength, columnLength)
addCombineTitleSpecialLast(titles, table)
addContentToCombineTitle(document, dc, table, titleLength, style=2)
addParagraph(document,"3、期末重要的债权投资","paragraph")
excelTableToWord("期末重要的债权投资",MODELPATH,style=2)

addTitle(document, "（十六）可供出售金融资产", 2, True)
addParagraph(document,"1、可供出售金融资产情况","paragraph")
df = pd.read_excel(MODELPATH, sheet_name="可供出售金融资产情况")
dc = df.to_dict("split")
titles = [["项  目", "期末数", "nan", "nan", "期初数", "nan", "nan"],
          ["nan", "账面余额", "减值准备", "账面价值","账面余额", "减值准备", "账面价值"]]
titleLength = len(titles)
rowLength = len(dc["index"]) + titleLength
columnLength = len(dc["columns"])
table = createBorderedTable(document, rowLength, columnLength)
addCombineTableTitle(table,titles)
addContentToCombineTitle(document, dc, table, titleLength, style=2)
addParagraph(document,"2、期末按成本计量的可供出售金融资产","paragraph")
df = pd.read_excel(MODELPATH, sheet_name="期末按成本计量的可供出售金融资产")
dc = df.to_dict("split")
titles = [["被投资单位", "账面余额", "nan", "nan", "nan", "减值准备", "nan", "nan", "nan","在被投资单位持股比例（%）","本期现金红利"],
          ["nan", "期初", "本期增加", "本期减少","期末", "期初", "本期增加", "本期减少","期末","nan", "nan"]]
titleLength = len(titles)
rowLength = len(dc["index"]) + titleLength
columnLength = len(dc["columns"])
table = createBorderedTable(document, rowLength, columnLength)
addCombineTitleSpecialReceivable(titles,table,[[1,9],[1,10]])
addContentToCombineTitle(document, dc, table, titleLength, style=2)
addParagraph(document,"3、期末按公允价值计量的可供出售金融资产","paragraph")
excelTableToWord("期末按公允价值计量的可供出售金融资产",MODELPATH,style=2)
addParagraph(document,"4、本期可供出售金融资产减值的变动情况","paragraph")
excelTableToWord("本期可供出售金融资产减值的变动情况",MODELPATH,style=2)
addParagraph(document,"5、可供出售权益工具年末公允价值严重下跌或非暂时性下跌但未计提减值准备的相关说明","paragraph")
excelTableToWord("可供出售权益工具严重下跌但未计提减值",MODELPATH,style=2)

addTitle(document, "（十七）其他债权投资", 2, True)
addParagraph(document,"1、明细情况","paragraph")
df = pd.read_excel(MODELPATH, sheet_name="其他债权投资期末数")
dc = df.to_dict("split")
titles = [["项  目", "期末数", "nan", "nan", "nan", "nan", "nan"],
          ["nan", "初始成本", "利息调整", "应计利息","公允价值变动","账面价值","减值准备"]]
titleLength = len(titles)
rowLength = len(dc["index"]) + titleLength
columnLength = len(dc["columns"])
table = createBorderedTable(document, rowLength, columnLength)
addCombineTableTitle(table,titles)
addContentToCombineTitle(document, dc, table, titleLength, style=2)
addParagraph(document,"（续）","paragraph")
df = pd.read_excel(MODELPATH, sheet_name="其他债权投资期初数")
dc = df.to_dict("split")
titles = [["项  目", "期初数", "nan", "nan", "nan", "nan", "nan"],
          ["nan", "初始成本", "利息调整", "应计利息","公允价值变动","账面价值","减值准备"]]
titleLength = len(titles)
rowLength = len(dc["index"]) + titleLength
columnLength = len(dc["columns"])
table = createBorderedTable(document, rowLength, columnLength)
addCombineTableTitle(table,titles)
addContentToCombineTitle(document, dc, table, titleLength, style=2)
addParagraph(document,"2、其他债权投资减值准备","paragraph")
df = pd.read_excel(MODELPATH, sheet_name="其他债权投资减值准备")
dc = df.to_dict("split")
titles = [["项目", "第一阶段", "第二阶段", "第三阶段", "合  计"],
          ["nan", "未来12个月预期信用损失", "整个存续期预期信用损失(未发生信用减值)", "整个存续期预期信用损失(已发生信用减值)","nan"]]
titleLength = len(titles)
rowLength = len(dc["index"]) + titleLength
columnLength = len(dc["columns"])
table = createBorderedTable(document, rowLength, columnLength)
addCombineTitleSpecialLast(titles, table)
addContentToCombineTitle(document, dc, table, titleLength, style=2)
addParagraph(document,"3、期末重要的其他债权投资","paragraph")
excelTableToWord("期末重要的其他债权投资",MODELPATH,style=2)

addTitle(document, "（十八）持有至到期投资", 2, True)
addParagraph(document,"1、明细情况","paragraph")
df = pd.read_excel(MODELPATH, sheet_name="持有至到期投资明细情况")
dc = df.to_dict("split")
titles = [["项  目", "期末数", "nan", "nan", "期初数", "nan", "nan"],
          ["nan", "账面余额", "减值准备", "账面价值","账面余额", "减值准备", "账面价值"]]
titleLength = len(titles)
rowLength = len(dc["index"]) + titleLength
columnLength = len(dc["columns"])
table = createBorderedTable(document, rowLength, columnLength)
addCombineTableTitle(table,titles)
addContentToCombineTitle(document, dc, table, titleLength, style=2)
addParagraph(document,"2、期末重要的持有至到期投资","paragraph")
excelTableToWord("期末重要的持有至到期投资",MODELPATH,style=2)

addTitle(document, "（十九）长期应收款", 2, True)
addParagraph(document,"1、明细情况","paragraph")
df = pd.read_excel(MODELPATH, sheet_name="长期应收款明细情况")
dc = df.to_dict("split")
titles = [["项  目", "期末数", "nan", "nan", "期初数", "nan", "nan","折现率区间"],
          ["nan", "账面余额", "减值准备", "账面价值","账面余额", "减值准备", "账面价值","nan"]]
titleLength = len(titles)
rowLength = len(dc["index"]) + titleLength
columnLength = len(dc["columns"])
table = createBorderedTable(document, rowLength, columnLength)
addCombineTableTitle(table,titles)
addContentToCombineTitle(document, dc, table, titleLength, style=2)
addParagraph(document,"2、减值准备计提情况","paragraph")
df = pd.read_excel(MODELPATH, sheet_name="长期应收款坏账准备变动情况新金融工具准则")
dc = df.to_dict("split")
titles = [["项目", "第一阶段", "第二阶段", "第三阶段", "合  计"],
          ["nan", "未来12个月预期信用损失", "整个存续期预期信用损失(未发生信用减值)", "整个存续期预期信用损失(已发生信用减值)","nan"]]
titleLength = len(titles)
rowLength = len(dc["index"]) + titleLength
columnLength = len(dc["columns"])
table = createBorderedTable(document, rowLength, columnLength)
addCombineTitleSpecialLast(titles, table)
addContentToCombineTitle(document, dc, table, titleLength, style=2)
addParagraph(document,"3、因金融资产转移而终止确认的长期应收款","paragraph")
excelTableToWord("因金融资产转移而终止确认的长期应收款",MODELPATH,style=2)
addParagraph(document,"4、转移长期应收款且继续涉入形成的资产、负债金额","paragraph")
excelTableToWord("转移长期应收款且继续涉入形成的资产负债金额",MODELPATH,style=2)

addTitle(document, "（二十）长期股权投资", 2, True)
addParagraph(document,"1、分类情况","paragraph")
excelTableToWord("长期股权投资分类情况",MODELPATH,style=2)

def addContent1(document,type,startYear,MODELPATH):
    addParagraph(document, "2、明细情况", "paragraph")
    df = pd.read_excel(MODELPATH, sheet_name="长期股权投资明细情况")
    dc = df.to_dict("split")
    titles = [["被投资单位","期初余额", "本期增减变动", "nan", "nan ", "nan", "nan","nan ", "nan", "nan","期末余额","减值准备期末余额"], ["nan", "nan", "追加投资", "减少投资", "权益法下确认的投资损益", "其他综合收益调整","其他权益变动","宣告发放现金股利或利润","计提减值准备","其他","nan","nan"]]
    titleLength = len(titles)
    rowLength = len(dc["index"]) + titleLength
    columnLength = len(dc["columns"])
    table = createBorderedTable(document,rowLength,columnLength)
    addCombineTitleSpecialReceivable(titles,table,[[1,10],[1,11]])
    addContentToCombineTitle(document, dc, table, titleLength, style=2)
addLandscapeContent(document,addContent1,type,startYear,MODELPATH)

addParagraph(document,"3、重要合营企业的主要财务信息","paragraph")
addParagraph(document,"本期数：","paragraph")
excelTableToWord("重要合营企业财务信息本期数",MODELPATH,style=2)
addParagraph(document,"上期数：","paragraph")
excelTableToWord("重要合营企业财务信息上期数",MODELPATH,style=2)

addParagraph(document,"4、重要联营企业的主要财务信息","paragraph")
addParagraph(document,"本期数：","paragraph")
excelTableToWord("重要联营企业财务信息本期数",MODELPATH,style=2)
addParagraph(document,"上期数：","paragraph")
excelTableToWord("重要联营企业财务信息上期数",MODELPATH,style=2)

addParagraph(document,"5、重要联营企业的主要财务信息","paragraph")
excelTableToWord("不重要合营企业和联营企业的汇总信息",MODELPATH,style=2)

addParagraph(document,"6、合营企业或联营企业发生的超额亏损","paragraph")
excelTableToWord("合营企业或联营企业发生的超额亏损",MODELPATH,style=2)

addTitle(document, "（二十一）其他权益工具投资", 2, True)
addParagraph(document,"1、明细情况","paragraph")
excelTableToWord("其他权益工具投资明细",MODELPATH,style=2)
addParagraph(document,"2、非交易性权益工具投资情况","paragraph")
excelTableToWord("非交易性权益工具投资情况",MODELPATH,style=2)

addTitle(document, "（二十二）其他非流动金融资产", 2, True)
excelTableToWord("其他非流动金融资产",MODELPATH,style=2)

addTitle(document, "（二十三）投资性房地产", 2, True)
addParagraph(document,"1、采用成本计量模式的投资性房地产","paragraph")
excelTableToWord("采用成本计量模式的投资性房地产",MODELPATH,style=2)
addParagraph(document,"2、采用公允价值计量模式的投资性房地产","paragraph")
excelTableToWord("采用公允价值计量模式的投资性房地产",MODELPATH,style=2)
addParagraph(document,"3、未办妥产权证书的投资性房地产金额及原因","paragraph")
excelTableToWord("未办妥产权证书的投资性房地产金额及原因",MODELPATH,style=2)

addTitle(document, "（二十四）固定资产", 2, True)
excelTableToWord("固定资产汇总",MODELPATH,style=2)
addParagraph(document,"1、固定资产","paragraph")
addParagraph(document,"（1）固定资产情况","paragraph")
excelTableToWord("固定资产情况",MODELPATH,style=2)
addParagraph(document,"（2）暂时闲置的固定资产情况","paragraph")
excelTableToWord("暂时闲置的固定资产情况",MODELPATH,style=2)
addParagraph(document,"（3）通过经营租赁租出的固定资产","paragraph")
excelTableToWord("通过经营租赁租出的固定资产",MODELPATH,style=2)
addParagraph(document,"（4）未办妥产权证书的固定资产情况","paragraph")
excelTableToWord("未办妥产权证书的固定资产情况",MODELPATH,style=2)
addParagraph(document,"2、固定资产清理","paragraph")
excelTableToWord("固定资产清理",MODELPATH,style=2)

addTitle(document, "（二十五）在建工程", 2, True)
excelTableToWord("在建工程汇总",MODELPATH,style=2)
addParagraph(document,"1、在建工程","paragraph")
addParagraph(document,"（1）在建工程情况","paragraph")
df = pd.read_excel(MODELPATH, sheet_name="在建工程情况")
dc = df.to_dict("split")
titles = [["项  目", "期末数", "nan", "nan", "期初数", "nan", "nan"],
          ["nan", "账面余额", "减值准备", "账面价值","账面余额", "减值准备", "账面价值"]]
titleLength = len(titles)
rowLength = len(dc["index"]) + titleLength
columnLength = len(dc["columns"])
table = createBorderedTable(document, rowLength, columnLength)
addCombineTableTitle(table,titles)
addContentToCombineTitle(document, dc, table, titleLength, style=2)

def addContent2(document,type,startYear,MODELPATH):
    addParagraph(document, "（2）重要在建工程项目本期变动情况", "paragraph")
    excelTableToWord("重要在建工程项目本期变动情况", MODELPATH, style=2)
addLandscapeContent(document,addContent2,type,startYear,MODELPATH)

addParagraph(document, "（3）本期计提在建工程减值准备情况", "paragraph")
excelTableToWord("本期计提在建工程减值准备情况", MODELPATH, style=2)

addParagraph(document,"2、工程物资","paragraph")
excelTableToWord("工程物资", MODELPATH, style=2)

addTitle(document, "（二十六）生产性生物资产", 2, True)
excelTableToWord("生产性生物资产", MODELPATH, style=2)

addTitle(document, "（二十七）油气资产", 2, True)
excelTableToWord("油气资产", MODELPATH, style=2)

addTitle(document, "（二十八）使用权资产", 2, True)
excelTableToWord("使用权资产", MODELPATH, style=2)

addTitle(document, "（二十九）无形资产", 2, True)
addParagraph(document, "1、明细情况", "paragraph")
excelTableToWord("无形资产", MODELPATH, style=2)
addParagraph(document, "2、未办妥产权证书的土地使用权情况", "paragraph")
excelTableToWord("未办妥产权证书的土地使用权情况", MODELPATH, style=2)

addTitle(document, "（三十）开发支出", 2, True)
df = pd.read_excel(MODELPATH, sheet_name="开发支出")
dc = df.to_dict("split")
titles = [["项  目", "期初数", "本期增加金额", "nan", "本期减少金额", "nan", "nan","期末数"],
          ["nan", "nan", "内部开发支出", "其他","确认为无形资产", "转入当期损益", "其他","nan"]]
titleLength = len(titles)
rowLength = len(dc["index"]) + titleLength
columnLength = len(dc["columns"])
table = createBorderedTable(document, rowLength, columnLength)
addCombineTitleSpecialReceivable(titles,table,[[1,7]])
addContentToCombineTitle(document, dc, table, titleLength, style=2)

addTitle(document, "（三十一）商誉", 2, True)
addParagraph(document, "1、商誉账面价值", "paragraph")
excelTableToWord("商誉账面价值", MODELPATH, style=2)
addParagraph(document, "2、	商誉减值准备", "paragraph")
excelTableToWord("商誉减值准备", MODELPATH, style=2)

addTitle(document, "（三十二）长期待摊费用", 2, True)
excelTableToWord("长期待摊费用", MODELPATH, style=2)

addTitle(document, "（三十三）递延所得税资产和递延所得税负债", 2, True)
addParagraph(document, "1、	未经抵销的递延所得税资产", "paragraph")
df = pd.read_excel(MODELPATH, sheet_name="未经抵销的递延所得税资产")
dc = df.to_dict("split")
titles = [["项  目", "期末数", "nan", "期初数", "nan"],
          ["nan", "可抵扣暂时性差异", "递延所得税资产", "可抵扣暂时性差异","递延所得税资产"]]
titleLength = len(titles)
rowLength = len(dc["index"]) + titleLength
columnLength = len(dc["columns"])
table = createBorderedTable(document, rowLength, columnLength)
addCombineTitleSpecialReceivable(titles,table,[[1,7]])
addContentToCombineTitle(document, dc, table, titleLength, style=2)
addParagraph(document, "2、	未经抵销的递延所得税负债", "paragraph")
df = pd.read_excel(MODELPATH, sheet_name="未经抵销的递延所得税负债")
dc = df.to_dict("split")
titles = [["项  目", "期末数", "nan", "期初数", "nan"],
          ["nan", "应纳税暂时性差异", "递延所得税负债", "应纳税暂时性差异","递延所得税负债"]]
titleLength = len(titles)
rowLength = len(dc["index"]) + titleLength
columnLength = len(dc["columns"])
table = createBorderedTable(document, rowLength, columnLength)
addCombineTitleSpecialReceivable(titles,table,[[1,7]])
addContentToCombineTitle(document, dc, table, titleLength, style=2)
addParagraph(document, "3、	未确认递延所得税资产明细", "paragraph")
excelTableToWord("未确认递延所得税资产明细", MODELPATH, style=2)
addParagraph(document, "4、	未确认递延所得税资产的可抵扣亏损将于以下年度到期", "paragraph")
excelTableToWord("未确认递延所得税资产的可抵扣亏损将于以下年度到期", MODELPATH, style=2)

addTitle(document, "（三十四）其他非流动资产", 2, True)
excelTableToWord("其他非流动资产", MODELPATH, style=2)

addTitle(document, "（三十五）短期借款", 2, True)
addParagraph(document, "1、	明细情况", "paragraph")
excelTableToWord("短期借款明细情况", MODELPATH, style=2)
addParagraph(document, "2、已逾期未偿还的短期借款情况", "paragraph")
excelTableToWord("已逾期未偿还的短期借款情况", MODELPATH, style=2)

addTitle(document, "（三十六）交易性金融负债", 2, True)
excelTableToWord("交易性金融负债", MODELPATH, style=2)

addTitle(document, "（三十七）以公允价值计量且其变动计入当期损益的金融负债", 2, True)
excelTableToWord("以公允价值计量且其变动计入当期损益的金融负债", MODELPATH, style=2)

addTitle(document, "（三十八）衍生金融负债", 2, True)
excelTableToWord("衍生金融负债", MODELPATH, style=2)

addTitle(document, "（三十九）应付票据", 2, True)
excelTableToWord("应付票据", MODELPATH, style=2)

addTitle(document, "（四十）应付账款", 2, True)
excelTableToWord("应付账款", MODELPATH, style=2)
addParagraph(document, "其中：账龄超过1年的重要应付账款", "paragraph")
excelTableToWord("账龄超过一年的重要应付账款", MODELPATH, style=2)

addTitle(document, "（四十一）预收款项", 2, True)
excelTableToWord("预收款项", MODELPATH, style=2)
addParagraph(document, "其中：账龄超过1年的重要预收款项", "paragraph")
excelTableToWord("账龄一年以上重要的预收款项", MODELPATH, style=2)

addTitle(document, "（四十二）合同负债", 2, True)
excelTableToWord("合同负债", MODELPATH, style=2)


addTitle(document, "（四十三）应付职工薪酬", 2, True)
addParagraph(document, "1、明细情况", "paragraph")
excelTableToWord("应付职工薪酬明细情况", MODELPATH, style=2)
addParagraph(document, "2、	短期薪酬列示", "paragraph")
excelTableToWord("短期薪酬列示", MODELPATH, style=2)
addParagraph(document, "3、设定提存计划列示", "paragraph")
excelTableToWord("设定提存计划列示", MODELPATH, style=2)

addTitle(document, "（四十四）应交税费", 2, True)
excelTableToWord("应交税费", MODELPATH, style=2)

addTitle(document, "（四十五）其他应付款", 2, True)
excelTableToWord("其他应付款汇总", MODELPATH, style=2)
addParagraph(document, "1、应付利息", "paragraph")
excelTableToWord("应付利息", MODELPATH, style=2)
addParagraph(document, "其中：重要的已逾期未支付的利息情况", "paragraph")
excelTableToWord("重要的已逾期未支付的利息情况", MODELPATH, style=2)
addParagraph(document, "2、	应付股利", "paragraph")
excelTableToWord("应付股利", MODELPATH, style=2)
addParagraph(document, "其中：账龄1年以上重要的应付股利", "paragraph")
excelTableToWord("账龄一年以上重要的应付股利", MODELPATH, style=2)
addParagraph(document, "3、	其他应付款项", "paragraph")
excelTableToWord("其他应付款项", MODELPATH, style=2)
addParagraph(document, "其中：账龄超过1年的重要其他应付款项", "paragraph")
excelTableToWord("账龄超过一年的重要其他应付款项", MODELPATH, style=2)

addTitle(document, "（四十六）持有待售负债", 2, True)
excelTableToWord("持有待售负债", MODELPATH, style=2)

addTitle(document, "（四十七）一年内到期的非流动负债", 2, True)
excelTableToWord("一年内到期的非流动负债", MODELPATH, style=2)

addTitle(document, "（四十八）其他流动负债", 2, True)
excelTableToWord("其他流动负债", MODELPATH, style=2)
def addContent3(document,type,startYear,MODELPATH):
    addParagraph(document, "其中：短期应付债券的增减变动", "paragraph")
    excelTableToWord("短期应付债券", MODELPATH, style=2)
addLandscapeContent(document,addContent3,type,startYear,MODELPATH)

addTitle(document, "（四十九）长期借款", 2, True)
excelTableToWord("长期借款", MODELPATH, style=2)

addTitle(document, "（五十）应付债券", 2, True)
addParagraph(document, "1、	明细情况", "paragraph")
excelTableToWord("应付债券", MODELPATH, style=2)
def addContent4(document,type,startYear,MODELPATH):
    addParagraph(document, "2、应付债券增减变动（不包括划分为金融负债的优先股、永续债等其他金融工具）", "paragraph")
    excelTableToWord("应付债券的增减变动", MODELPATH, style=2)
addLandscapeContent(document,addContent4,type,startYear,MODELPATH)

addTitle(document, "（五十一）优先股、永续债等金融工具", 2, True)
addParagraph(document, "1、	期末发行在外的优先股、永续债等金融工具情况", "paragraph")
excelTableToWord("期末发行在外的优先股永续债等金融工具情况", MODELPATH, style=2)
addParagraph(document, "2、	发行在外的优先股、永续债等金融工具变动情况", "paragraph")
df = pd.read_excel(MODELPATH, sheet_name="发行在外的优先股永续债等金融工具变动情况")
dc = df.to_dict("split")
titles = [["发行在外的金融工具", "期初余额", "nan", "本期增加", "nan", "本期减少", "nan", "期末余额", "nan"],
          ["nan", "数量", "账面价值","数量", "账面价值","数量", "账面价值","数量", "账面价值"]]
titleLength = len(titles)
rowLength = len(dc["index"]) + titleLength
columnLength = len(dc["columns"])
table = createBorderedTable(document, rowLength, columnLength)
addCombineTitleSpecialReceivable(titles,table,[[1,7]])
addContentToCombineTitle(document, dc, table, titleLength, style=2)
addParagraph(document, "3、	归属于权益工具持有者的信息", "paragraph")
excelTableToWord("归属于权益工具持有者的信息", MODELPATH, style=2)

addTitle(document, "（五十二）租赁负债", 2, True)
excelTableToWord("租赁负债", MODELPATH, style=2)

addTitle(document, "（五十三）长期应付款", 2, True)
excelTableToWord("长期应付款汇总", MODELPATH, style=2)
addParagraph(document, "1、	长期应付款", "paragraph")
excelTableToWord("长期应付款", MODELPATH, style=2)
addParagraph(document, "2、	专项应付款", "paragraph")
excelTableToWord("专项应付款", MODELPATH, style=2)

addTitle(document, "（五十四）长期应付职工薪酬", 2, True)
addParagraph(document, "1、	明细情况", "paragraph")
excelTableToWord("长期应付职工薪酬明细情况", MODELPATH, style=2)
addParagraph(document, "2、	设定受益计划变动情况", "paragraph")
addParagraph(document, "（1）设定受益计划义务现值", "paragraph")
excelTableToWord("设定受益计划义务现值", MODELPATH, style=2)
addParagraph(document, "（2）计划资产", "paragraph")
excelTableToWord("计划资产", MODELPATH, style=2)
addParagraph(document, "（3）设定受益计划净负债（净资产）", "paragraph")
excelTableToWord("设定受益计划净负债", MODELPATH, style=2)

addTitle(document, "（五十五）预计负债", 2, True)
excelTableToWord("预计负债", MODELPATH, style=2)

addTitle(document, "（五十六）递延收益", 2, True)
excelTableToWord("递延收益", MODELPATH, style=2)
addParagraph(document, "其中，涉及政府补助的项目：", "paragraph")
excelTableToWord("递延收益中政府补助项目", MODELPATH, style=2)

addTitle(document, "（五十七）其他非流动负债", 2, True)
excelTableToWord("其他非流动负债", MODELPATH, style=2)

addTitle(document, "（五十八）实收资本", 2, True)
df = pd.read_excel(MODELPATH, sheet_name="实收资本")
dc = df.to_dict("split")
titles = [["投资者名称", "期初余额", "nan", "本期增加", "本期减少",  "期末余额", "nan"],
          ["nan", "投资金额", "所占比例(%)","nan", "nan","投资金额", "所占比例(%)"]]
titleLength = len(titles)
rowLength = len(dc["index"]) + titleLength
columnLength = len(dc["columns"])
table = createBorderedTable(document, rowLength, columnLength)
addCombineTitleSpecialReceivable(titles,table,[[1,3],[1,4]])
addContentToCombineTitle(document, dc, table, titleLength, style=2)

addTitle(document, "（五十八）股本", 2, True)
df = pd.read_excel(MODELPATH, sheet_name="股本")
dc = df.to_dict("split")
titles = [["项  目", "期初数", "本期增减变动（减少以“—”表示）", "nan", "nan",  "nan", "nan","期末数"],
          ["nan", "nan", "发行新股","送股", "公积金转股","其他", "小计","nan"]]
titleLength = len(titles)
rowLength = len(dc["index"]) + titleLength
columnLength = len(dc["columns"])
table = createBorderedTable(document, rowLength, columnLength)
addCombineTitleSpecialReceivable(titles,table,[[1,7]])
addContentToCombineTitle(document, dc, table, titleLength, style=2)

addTitle(document, "（五十九）其他权益工具", 2, True)
df = pd.read_excel(MODELPATH, sheet_name="其他权益工具")
dc = df.to_dict("split")
titles = [["项  目", "期初余额", "nan", "本期增加", "nan", "本期减少", "nan", "期末余额", "nan"],
          ["nan", "数量", "账面价值","数量", "账面价值","数量", "账面价值","数量", "账面价值"]]
titleLength = len(titles)
rowLength = len(dc["index"]) + titleLength
columnLength = len(dc["columns"])
table = createBorderedTable(document, rowLength, columnLength)
addCombineTitleSpecialReceivable(titles,table,[[1,7]])
addContentToCombineTitle(document, dc, table, titleLength, style=2)

addTitle(document, "（六十）资本公积", 2, True)
excelTableToWord("资本公积", MODELPATH, style=2)

addTitle(document, "（六十一）其他综合收益", 2, True)
df = pd.read_excel(MODELPATH, sheet_name="其他综合收益")
dc = df.to_dict("split")
titles = [["项  目", "期初数", "本期发生额", "nan", "nan", "nan", "nan", "nan", "期末数"],
          ["nan", "nan", "本期所得税前发生额","减：前期计入其他综合收益当期转入损益", "减：前期计入其他综合收益当期转入留存收益","减：所得税费用", "税后归属于母公司","税后归属于少数股东", "nan"]]
titleLength = len(titles)
rowLength = len(dc["index"]) + titleLength
columnLength = len(dc["columns"])
table = createBorderedTable(document, rowLength, columnLength)
addCombineTitleSpecialReceivable(titles,table,[[1,8]])
addContentToCombineTitle(document, dc, table, titleLength, style=2)

addTitle(document, "（六十一）专项储备", 2, True)
excelTableToWord("专项储备", MODELPATH, style=2)

addTitle(document, "（六十二）盈余公积", 2, True)
excelTableToWord("盈余公积", MODELPATH, style=2)

addTitle(document, "（六十三）未分配利润", 2, True)
excelTableToWord("未分配利润", MODELPATH, style=2)

addTitle(document, "（六十四）营业收入、营业成本", 2, True)
df = pd.read_excel(MODELPATH, sheet_name="营业收入与营业成本")
dc = df.to_dict("split")
titles = [["项  目", "本期数", "nan", "上年同期数", "nan"],
          ["nan", "收入", "成本","收入", "成本"]]
titleLength = len(titles)
rowLength = len(dc["index"]) + titleLength
columnLength = len(dc["columns"])
table = createBorderedTable(document, rowLength, columnLength)
addCombineTableTitle(table,titles)
addContentToCombineTitle(document, dc, table, titleLength, style=2)
addParagraph(document, "1、主营业务收入和主营业务成本", "paragraph")
df = pd.read_excel(MODELPATH, sheet_name="主营业务收入与主营业务成本")
dc = df.to_dict("split")
titles = [["项  目", "本期数", "nan", "上年同期数", "nan"],
          ["nan", "收入", "成本","收入", "成本"]]
titleLength = len(titles)
rowLength = len(dc["index"]) + titleLength
columnLength = len(dc["columns"])
table = createBorderedTable(document, rowLength, columnLength)
addCombineTableTitle(table,titles)
addContentToCombineTitle(document, dc, table, titleLength, style=2)
addParagraph(document, "2、其他业务收入和其他业务成本", "paragraph")
df = pd.read_excel(MODELPATH, sheet_name="其他业务收入与其他业务成本")
dc = df.to_dict("split")
titles = [["项  目", "本期数", "nan", "上年同期数", "nan"],
          ["nan", "收入", "成本","收入", "成本"]]
titleLength = len(titles)
rowLength = len(dc["index"]) + titleLength
columnLength = len(dc["columns"])
table = createBorderedTable(document, rowLength, columnLength)
addCombineTableTitle(table,titles)
addContentToCombineTitle(document, dc, table, titleLength, style=2)

addTitle(document, "（六十五）税金及附加", 2, True)
excelTableToWord("税金及附加", MODELPATH, style=2)

addTitle(document, "（六十五）销售费用", 2, True)
excelTableToWord("销售费用", MODELPATH, style=2)

addTitle(document, "（六十六）管理费用", 2, True)
excelTableToWord("管理费用", MODELPATH, style=2)

addTitle(document, "（六十六）研发费用", 2, True)
excelTableToWord("研发费用", MODELPATH, style=2)

addTitle(document, "（六十七）财务费用", 2, True)
excelTableToWord("财务费用", MODELPATH, style=2)

addTitle(document, "（六十一）其他收益", 2, True)
excelTableToWord("其他收益", MODELPATH, style=2)

addTitle(document, "（六十七）投资收益", 2, True)
excelTableToWord("投资收益", MODELPATH, style=2)

addTitle(document, "（六十八）净敞口套期收益", 2, True)
excelTableToWord("净敞口套期收益", MODELPATH, style=2)

addTitle(document, "（六十九）公允价值变动收益", 2, True)
excelTableToWord("公允价值变动损益", MODELPATH, style=2)

addTitle(document, "（七十）信用减值损失", 2, True)
excelTableToWord("信用减值损失", MODELPATH, style=2)

addTitle(document, "（七十一）资产减值损失", 2, True)
excelTableToWord("资产减值损失", MODELPATH, style=2)

addTitle(document, "（七十二）资产处置收益", 2, True)
excelTableToWord("资产处置收益", MODELPATH, style=2)

addTitle(document, "（七十三）营业外收入", 2, True)
excelTableToWord("营业外收入", MODELPATH, style=2)

addTitle(document, "（七十四）营业外支出", 2, True)
excelTableToWord("营业外支出", MODELPATH, style=2)

addTitle(document, "（七十五）所得税费用", 2, True)
addParagraph(document, "1、明细情况", "paragraph")
excelTableToWord("所得税费用", MODELPATH, style=2)
addParagraph(document, "2、会计利润与所得税费用调整过程", "paragraph")
excelTableToWord("会计利润与所得税费用调整过程", MODELPATH, style=2)

addTitle(document, "（七十六）归属于母公司所有者的其他综合收益", 2, True)
addParagraph(document, "详见附注六、54。", "paragraph")

addTitle(document, "（七十七）每股收益", 2, True)
addTitle(document, "（七十八）非货币性资产交换", 2, True)
addTitle(document, "（七十九）股份支付", 2, True)
addTitle(document, "（八十）债务重组", 2, True)
addTitle(document, "（八十一）借款费用", 2, True)
addTitle(document, "（八十二）外币折算", 2, True)
addTitle(document, "（八十三）租赁", 2, True)
addTitle(document, "（八十四）终止经营", 2, True)
addTitle(document, "（八十五）分部信息", 2, True)
addTitle(document, "（八十六）合并现金流量表", 2, True)
addTitle(document, "（八十七）外币货币性项目", 2, True)
addTitle(document, "（八十八）所有权和使用权受到限制的资产", 2, True)
addTitle(document, "九、或有事项", 1, False)
addTitle(document, "十、资产负债表日后事项", 1, False)
addTitle(document, "十一、关联方关系及其交易", 1, False)
addTitle(document, "（一）母公司基本情况", 2, True)
addTitle(document, "（二）子公司情况", 2, True)
addTitle(document, "（三）合营企业及联营企业情况", 2, True)
addTitle(document, "（四）其他关联方", 2, True)
addTitle(document, "（五）关联方交易", 2, True)
addTitle(document, "（五）关联方交易", 2, True)

document.save("noteappended.docx")