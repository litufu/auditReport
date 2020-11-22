# -*- coding: UTF-8 -*-

from docx import Document
import pandas as pd
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.enum.section import WD_ORIENTATION,WD_SECTION_START
from docx.shared import Pt,Cm

from project.data import report_params, standardChange,notes_params,noteAppend
from project.settings import setStyle
from project.utils import checkLeftSpace, addTitle, addCombineTableTitle,addContentToCombineTitle, addParagraph,createBorderedTable,addTable,setCell,addLandscapeContent


# 国资报表
# 已实施新准则
# 未实施新准则
# 首次执行新准则
# 报告日期
reportDate = report_params["reportDate"]
companyName = report_params["companyName"]
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
def addTableTitle(name):
    paragraph = document.add_paragraph()
    paragraph_format = paragraph.paragraph_format
    paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    paragraph.add_run(name,style="first")

def addFirstLine(document,companyName,reportDate,currencyUnit):
    table = document.add_table(1, 3)
    table.rows[0].height_rule  =WD_ROW_HEIGHT_RULE.EXACTLY
    table.rows[0].height = Cm(0.8)
    table.cell(0, 0).width = Cm(20)
    table.cell(0, 1).width = Cm(10)
    table.cell(0, 2).width = Cm(10)
    setCell(table.cell(0, 0), "编制单位：{}".format(companyName), WD_PARAGRAPH_ALIGNMENT.LEFT, toFloat=True, style="tableSmallCharacter")
    setCell(table.cell(0, 1), reportDate, WD_PARAGRAPH_ALIGNMENT.CENTER, toFloat=True, style="tableSmallCharacter")
    setCell(table.cell(0, 2), "单位：{}".format(currencyUnit), WD_PARAGRAPH_ALIGNMENT.RIGHT, toFloat=True, style="tableSmallCharacter")

balanceTitles = ["项            目","行次","年末余额","年初余额","注释号"]
assetsRecords = [
        {"name":"项            目","type":"center","origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"流动资产：","type":0,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"货币资金","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"△结算备付金","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"△拆出资金","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"☆交易性金融资产","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"以公允价值计量且其变动计入当期损益的金融资产","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"衍生金融资产","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"应收票据","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"应收账款","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"☆应收款项融资","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"预付款项","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"△应收保费","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"△应收分保账款","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"△应收分保合同准备金","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"其他应收款","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"△买入返售金融资产","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"存货","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"其中：原材料","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"库存商品(产成品)","type":3,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"☆合同资产","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"持有待售资产","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"一年内到期的非流动资产","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"其他流动资产","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"流动资产合计","type":"center","origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"非流动资产：","type":0,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"△发放贷款和垫款","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"☆债权投资","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"可供出售金融资产","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"☆其他债权投资","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"持有至到期投资","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"长期应收款","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"长期股权投资","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"☆其他权益工具投资","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"☆其他非流动金融资产","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"投资性房地产","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"固定资产","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"在建工程","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"生产性生物资产","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"油气资产","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"☆使用权资产","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"无形资产","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"开发支出","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"商誉","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"长期待摊费用","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"递延所得税资产","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"其他非流动资产","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"其中：特准储备物资","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"非流动资产合计","type":"center","origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"资  产  总  计","type":"center","origin":[],"endDate":0.00,"startDate":0.00},
]

def setHeader(table,titles):
    table.cell(0, 0).width = Cm(13)
    table.cell(0, 1).width = Cm(2)
    table.cell(0, 2).width = Cm(6)
    table.cell(0, 3).width = Cm(6)
    table.cell(0, 4).width = Cm(3)
    setCell(table.cell(0, 0), titles[0], WD_PARAGRAPH_ALIGNMENT.CENTER, False, "tableSmallCharacter")
    setCell(table.cell(0, 1), titles[1], WD_PARAGRAPH_ALIGNMENT.CENTER, False, "tableSmallCharacter")
    setCell(table.cell(0, 2), titles[2], WD_PARAGRAPH_ALIGNMENT.CENTER, False, "tableSmallCharacter")
    setCell(table.cell(0, 3), titles[3], WD_PARAGRAPH_ALIGNMENT.CENTER, False, "tableSmallCharacter")
    setCell(table.cell(0, 4), titles[4], WD_PARAGRAPH_ALIGNMENT.CENTER, False, "tableSmallCharacter")

def getAlignAndText(type,originText):
    if type=="center":
        return originText,WD_PARAGRAPH_ALIGNMENT.CENTER
    else:
        return  " "*4*type+originText,WD_PARAGRAPH_ALIGNMENT.LEFT

def add_last_line():
    table = document.add_table(1, 3)
    table.cell(0, 0).width = Cm(10)
    table.cell(0, 1).width = Cm(13)
    table.cell(0, 2).width = Cm(10)
    setCell(table.cell(0, 0), "法定代表人：", WD_PARAGRAPH_ALIGNMENT.LEFT, toFloat=True,
            style="tableSmallCharacter")
    setCell(table.cell(0, 1), "主管会计工作的负责人：", WD_PARAGRAPH_ALIGNMENT.LEFT, toFloat=True, style="tableSmallCharacter")
    setCell(table.cell(0, 2), "会计机构负责人：", WD_PARAGRAPH_ALIGNMENT.LEFT, toFloat=True,
            style="tableSmallCharacter")

liabilitiesRecords = [
        {"name":"项            目","type":"center","origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"流动负债：","type":0,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"短期借款","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"△向中央银行借款","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"△拆入资金","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"☆交易性金融负债","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"以公允价值计量且其变动计入当期损益的金融负债","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"衍生金融负债","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"应付票据","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"应付账款","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"预收款项","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"☆合同负债","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"△卖出回购金融资产款","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"△吸收存款及同业存放","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"△代理买卖证券款","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"△代理承销证券款","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"应付职工薪酬","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"其中：应付工资","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"应付福利费","type":3,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"#其中：职工奖励及福利基金)","type":4,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"应交税费","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"其中：应交税金","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"其他应付款","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"△应付手续费及佣金","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"△应付分保账款","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"持有待售负债","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"一年内到期的非流动负债","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"其他流动负债","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"流动负债合计","type":"center","origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"非流动负债：","type":0,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"△保险合同准备金","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"长期借款","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"应付债券","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"其中：优先股","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"永续债","type":3,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"☆租赁负债","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"长期应付款","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"长期应付职工薪酬","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"预计负债","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"递延收益","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"递延所得税负债","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"其他非流动负债","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"其中：特准储备基金","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"非流动负债合计","type":"center","origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"负 债 合 计","type":"center","origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"所有者权益（或股东权益）：","type":0,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"实收资本（或股本）","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"国家资本","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"国有法人资本","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"集体资本","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"民营资本","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"外商资本","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"#减：已归还投资","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"实收资本（或股本）净额","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"其他权益工具","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"其中：优先股","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"永续债","type":3,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"资本公积","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"减：库存股","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"其他综合收益","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"其中：外币报表折算差额","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"专项储备","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"盈余公积","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"其中：法定公积金","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"任意公积金","type":3,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"#储备基金","type":3,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"#企业发展基金","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"#利润归还投资","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"△一般风险准备","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"未分配利润","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"归属于母公司所有者权益（或股东权益）合计","type":"center","origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"*少数股东权益","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"所有者权益（或股东权益）合计","type":"center","origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"负债和所有者权益（或股东权益）总计","type":"center","origin":[],"endDate":0.00,"startDate":0.00},
]
profitTitles = ["项            目","行次","本年金额","上年金额","注释号"]
profitRecords = [
        {"name":"项            目","type":"center","origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"一、营业总收入","type":0,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"其中：营业收入","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"△利息收入","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"△已赚保费","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"△手续费及佣金收入","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"二、营业总成本","type":0,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"其中：营业成本","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"△利息支出","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"△手续费及佣金支出","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"△退保金","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"△赔付支出净额","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"△提取保险责任准备金净额","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"△保单红利支出","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"△分保费用","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"税金及附加","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"销售费用","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"管理费用","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"其中：党建工作经费","type":3,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"研发费用","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"财务费用","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"其中：利息费用","type":3,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"利息收入","type":4,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"汇兑净损失（净收益以“-”号填列）","type":4,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"其他","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"加：其他收益","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"投资收益（损失以“-”号填列）","type":3,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"其中：对联营企业和合营企业的投资收益","type":4,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"☆以摊余成本计量的金融资产终止确认收益","type":5,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"△汇兑收益（损失以“-”号填列）","type":3,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"☆净敞口套期收益（损失以“-”号填列)","type":3,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"公允价值变动收益（损失以“-”号填列）","type":3,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"☆信用减值损失（损失以“-”号填列）","type":3,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"资产减值损失（损失以“-”号填列）","type":3,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"资产处置收益（损失以“-”号填列）","type":3,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"三、营业利润（亏损以“－”号填列）","type":0,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"加：营业外收入","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"其中：政府补助","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"减：营业外支出","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"四、利润总额（亏损总额以“－”号填列）","type":0,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"减：所得税费用","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"五、净利润（净亏损以“－”号填列）","type":0,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"（一）按所有权归属分类:","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"归属于母公司所有者的净利润","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"*少数股东损益","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"（二）按经营持续性分类:","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"持续经营净利润","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"终止经营净利润","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"六、其他综合收益的税后净额","type":0,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"归属于母公司所有者的其他综合收益的税后净额","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"一）不能重分类进损益的其他综合收益","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"1.重新计量设定受益计划变动额","type":3,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"2.权益法下不能转损益的其他综合收益","type":3,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"☆3.其他权益工具投资公允价值变动","type":3,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"☆4.企业自身信用风险公允价值变动","type":3,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"5.其他","type":3,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"（二）将重分类进损益的其他综合收益","type":2,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"1.权益法下可转损益的其他综合收益","type":3,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"☆2.其他债权投资公允价值变动","type":3,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"3.可供出售金融资产公允价值变动损益","type":3,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"☆4.金融资产重分类计入其他综合收益的金额","type":3,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"5.持有至到期投资重分类为可供出售金融资产损益","type":3,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"☆6.其他债权投资信用减值准备","type":3,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"7.现金流量套期储备（现金流量套期损益的有效部分）","type":3,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"8.外币财务报表折算差额","type":3,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"9.其他","type":3,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"*归属于少数股东的其他综合收益的税后净额","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"七、综合收益总额","type":0,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"归属于母公司所有者的综合收益总额","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"*归属于少数股东的综合收益总额","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"八、每股收益：","type":0,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"基本每股收益","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
        {"name":"稀释每股收益","type":1,"origin":[],"endDate":0.00,"startDate":0.00},
]

cashRecords = [
    {"name": "项              目", "type": "center", "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "一、经营活动产生的现金流量：", "type": 0, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "销售商品、提供劳务收到的现金", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "△客户存款和同业存放款项净增加额", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "△向中央银行借款净增加额", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "△向其他金融机构拆入资金净增加额", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "△收到原保险合同保费取得的现金", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "△收到再保业务现金净额", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "△保户储金及投资款净增加额", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "△处置以公允价值计量且其变动计入当期损益的金融资产净增加额", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "△收取利息、手续费及佣金的现金", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "△拆入资金净增加额", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "△回购业务资金净增加额", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "△代理买卖证券收到的现金净额", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "收到的税费返还", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "收到其他与经营活动有关的现金", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "经营活动现金流入小计", "type": "center", "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "购买商品、接受劳务支付的现金", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "△客户贷款及垫款净增加额", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "△存放中央银行和同业款项净增加额", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "△支付原保险合同赔付款项的现金", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "△拆出资金净增加额", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "△支付利息、手续费及佣金的现金", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "△支付保单红利的现金", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "支付给职工及为职工支付的现金", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "支付的各项税费", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "支付其他与经营活动有关的现金", "type":1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "经营活动现金流出小计", "type": "center", "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "经营活动产生的现金流量净额", "type": "center", "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "二、投资活动产生的现金流量：", "type": 0, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "收回投资收到的现金", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "取得投资收益收到的现金", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "处置固定资产、无形资产和其他长期资产收回的现金净额", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "处置子公司及其他营业单位收到的现金净额", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "收到其他与投资活动有关的现金", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "投资活动现金流入小计", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "购建固定资产、无形资产和其他长期资产支付的现金", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "投资支付的现金", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "△质押贷款净增加额", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "取得子公司及其他营业单位支付的现金净额", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "支付其他与投资活动有关的现金", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "投资活动现金流出小计", "type": "center", "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "投资活动产生的现金流量净额", "type": "center", "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "三、筹资活动产生的现金流量：", "type": 0, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "吸收投资收到的现金", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "其中：子公司吸收少数股东投资收到的现金", "type": 2, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "取得借款收到的现金", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "△发行债券收到的现金", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "收到其他与筹资活动有关的现金", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "筹资活动现金流入小计", "type": "center", "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "偿还债务支付的现金", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "分配股利、利润或偿付利息支付的现金", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "其中：子公司支付给少数股东的股利、利润", "type": 2, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "支付其他与筹资活动有关的现金", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "筹资活动现金流出小计", "type": "center", "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "筹资活动产生的现金流量净额", "type": "center", "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "四、汇率变动对现金及现金等价物的影响", "type": 0, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "五、现金及现金等价物净增加额", "type": 0, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "加：期初现金及现金等价物余额", "type": 1, "origin": [], "endDate": 0.00, "startDate": 0.00},
    {"name": "六、期末现金及现金等价物余额", "type": 0, "origin": [], "endDate": 0.00, "startDate": 0.00},
]

ownerRecords = [
        {"name":"项            目","type":"center","origin":[],"paidInCapital":0.00,"preferedStock":0.00,"perpetualDebt":0.00,"otherEquityInstruments":0.00,"capitalReserve":0.00,"treasuryStock":0.00,"otherComprehensiveIncome":0.00,"specialReserve":0.00,"surplusReserve":0.00,"generalRiskReserve":0.00,"undistributedProfit":0.00,"subtotal":0.00,"minorityInterests":0.00,"totalOwnerEquity":0.00},
        {"name":"一、上年年末余额","type":0,"origin":[],"paidInCapital":0.00,"preferedStock":0.00,"perpetualDebt":0.00,"otherEquityInstruments":0.00,"capitalReserve":0.00,"treasuryStock":0.00,"otherComprehensiveIncome":0.00,"specialReserve":0.00,"surplusReserve":0.00,"generalRiskReserve":0.00,"undistributedProfit":0.00,"subtotal":0.00,"minorityInterests":0.00,"totalOwnerEquity":0.00},
        {"name":"加：会计政策变更","type":1,"origin":[],"paidInCapital":0.00,"preferedStock":0.00,"perpetualDebt":0.00,"otherEquityInstruments":0.00,"capitalReserve":0.00,"treasuryStock":0.00,"otherComprehensiveIncome":0.00,"specialReserve":0.00,"surplusReserve":0.00,"generalRiskReserve":0.00,"undistributedProfit":0.00,"subtotal":0.00,"minorityInterests":0.00,"totalOwnerEquity":0.00},
        {"name":"前期差错更正","type":2,"origin":[],"paidInCapital":0.00,"preferedStock":0.00,"perpetualDebt":0.00,"otherEquityInstruments":0.00,"capitalReserve":0.00,"treasuryStock":0.00,"otherComprehensiveIncome":0.00,"specialReserve":0.00,"surplusReserve":0.00,"generalRiskReserve":0.00,"undistributedProfit":0.00,"subtotal":0.00,"minorityInterests":0.00,"totalOwnerEquity":0.00},
        {"name":"其他","type":2,"origin":[],"paidInCapital":0.00,"preferedStock":0.00,"perpetualDebt":0.00,"otherEquityInstruments":0.00,"capitalReserve":0.00,"treasuryStock":0.00,"otherComprehensiveIncome":0.00,"specialReserve":0.00,"surplusReserve":0.00,"generalRiskReserve":0.00,"undistributedProfit":0.00,"subtotal":0.00,"minorityInterests":0.00,"totalOwnerEquity":0.00},
        {"name":"二、本年年初余额","type":0,"origin":[],"paidInCapital":0.00,"preferedStock":0.00,"perpetualDebt":0.00,"otherEquityInstruments":0.00,"capitalReserve":0.00,"treasuryStock":0.00,"otherComprehensiveIncome":0.00,"specialReserve":0.00,"surplusReserve":0.00,"generalRiskReserve":0.00,"undistributedProfit":0.00,"subtotal":0.00,"minorityInterests":0.00,"totalOwnerEquity":0.00},
        {"name":"三、本年增减变动金额（减少以“-”号填列）","type":0,"origin":[],"paidInCapital":0.00,"preferedStock":0.00,"perpetualDebt":0.00,"otherEquityInstruments":0.00,"capitalReserve":0.00,"treasuryStock":0.00,"otherComprehensiveIncome":0.00,"specialReserve":0.00,"surplusReserve":0.00,"generalRiskReserve":0.00,"undistributedProfit":0.00,"subtotal":0.00,"minorityInterests":0.00,"totalOwnerEquity":0.00},
        {"name":"（一）综合收益总额","type":1,"origin":[],"paidInCapital":0.00,"preferedStock":0.00,"perpetualDebt":0.00,"otherEquityInstruments":0.00,"capitalReserve":0.00,"treasuryStock":0.00,"otherComprehensiveIncome":0.00,"specialReserve":0.00,"surplusReserve":0.00,"generalRiskReserve":0.00,"undistributedProfit":0.00,"subtotal":0.00,"minorityInterests":0.00,"totalOwnerEquity":0.00},
        {"name":"（二）所有者投入和减少资本","type":1,"origin":[],"paidInCapital":0.00,"preferedStock":0.00,"perpetualDebt":0.00,"otherEquityInstruments":0.00,"capitalReserve":0.00,"treasuryStock":0.00,"otherComprehensiveIncome":0.00,"specialReserve":0.00,"surplusReserve":0.00,"generalRiskReserve":0.00,"undistributedProfit":0.00,"subtotal":0.00,"minorityInterests":0.00,"totalOwnerEquity":0.00},
        {"name":"1.所有者投入资本","type":2,"origin":[],"paidInCapital":0.00,"preferedStock":0.00,"perpetualDebt":0.00,"otherEquityInstruments":0.00,"capitalReserve":0.00,"treasuryStock":0.00,"otherComprehensiveIncome":0.00,"specialReserve":0.00,"surplusReserve":0.00,"generalRiskReserve":0.00,"undistributedProfit":0.00,"subtotal":0.00,"minorityInterests":0.00,"totalOwnerEquity":0.00},
        {"name":"2.其他权益工具持有者投入资本","type":2,"origin":[],"paidInCapital":0.00,"preferedStock":0.00,"perpetualDebt":0.00,"otherEquityInstruments":0.00,"capitalReserve":0.00,"treasuryStock":0.00,"otherComprehensiveIncome":0.00,"specialReserve":0.00,"surplusReserve":0.00,"generalRiskReserve":0.00,"undistributedProfit":0.00,"subtotal":0.00,"minorityInterests":0.00,"totalOwnerEquity":0.00},
        {"name":"3.股份支付计入所有者权益的金额","type":2,"origin":[],"paidInCapital":0.00,"preferedStock":0.00,"perpetualDebt":0.00,"otherEquityInstruments":0.00,"capitalReserve":0.00,"treasuryStock":0.00,"otherComprehensiveIncome":0.00,"specialReserve":0.00,"surplusReserve":0.00,"generalRiskReserve":0.00,"undistributedProfit":0.00,"subtotal":0.00,"minorityInterests":0.00,"totalOwnerEquity":0.00},
        {"name":"4.其他","type":2,"origin":[],"paidInCapital":0.00,"preferedStock":0.00,"perpetualDebt":0.00,"otherEquityInstruments":0.00,"capitalReserve":0.00,"treasuryStock":0.00,"otherComprehensiveIncome":0.00,"specialReserve":0.00,"surplusReserve":0.00,"generalRiskReserve":0.00,"undistributedProfit":0.00,"subtotal":0.00,"minorityInterests":0.00,"totalOwnerEquity":0.00},
        {"name":"（三）专项储备提取和使用","type":1,"origin":[],"paidInCapital":0.00,"preferedStock":0.00,"perpetualDebt":0.00,"otherEquityInstruments":0.00,"capitalReserve":0.00,"treasuryStock":0.00,"otherComprehensiveIncome":0.00,"specialReserve":0.00,"surplusReserve":0.00,"generalRiskReserve":0.00,"undistributedProfit":0.00,"subtotal":0.00,"minorityInterests":0.00,"totalOwnerEquity":0.00},
        {"name":"1.提取专项储备","type":2,"origin":[],"paidInCapital":0.00,"preferedStock":0.00,"perpetualDebt":0.00,"otherEquityInstruments":0.00,"capitalReserve":0.00,"treasuryStock":0.00,"otherComprehensiveIncome":0.00,"specialReserve":0.00,"surplusReserve":0.00,"generalRiskReserve":0.00,"undistributedProfit":0.00,"subtotal":0.00,"minorityInterests":0.00,"totalOwnerEquity":0.00},
        {"name":"2.使用专项储备","type":2,"origin":[],"paidInCapital":0.00,"preferedStock":0.00,"perpetualDebt":0.00,"otherEquityInstruments":0.00,"capitalReserve":0.00,"treasuryStock":0.00,"otherComprehensiveIncome":0.00,"specialReserve":0.00,"surplusReserve":0.00,"generalRiskReserve":0.00,"undistributedProfit":0.00,"subtotal":0.00,"minorityInterests":0.00,"totalOwnerEquity":0.00},
        {"name":"（四）利润分配","type":1,"origin":[],"paidInCapital":0.00,"preferedStock":0.00,"perpetualDebt":0.00,"otherEquityInstruments":0.00,"capitalReserve":0.00,"treasuryStock":0.00,"otherComprehensiveIncome":0.00,"specialReserve":0.00,"surplusReserve":0.00,"generalRiskReserve":0.00,"undistributedProfit":0.00,"subtotal":0.00,"minorityInterests":0.00,"totalOwnerEquity":0.00},
        {"name":"1.提取盈余公积","type":2,"origin":[],"paidInCapital":0.00,"preferedStock":0.00,"perpetualDebt":0.00,"otherEquityInstruments":0.00,"capitalReserve":0.00,"treasuryStock":0.00,"otherComprehensiveIncome":0.00,"specialReserve":0.00,"surplusReserve":0.00,"generalRiskReserve":0.00,"undistributedProfit":0.00,"subtotal":0.00,"minorityInterests":0.00,"totalOwnerEquity":0.00},
        {"name":"其中：法定公积金","type":3,"origin":[],"paidInCapital":0.00,"preferedStock":0.00,"perpetualDebt":0.00,"otherEquityInstruments":0.00,"capitalReserve":0.00,"treasuryStock":0.00,"otherComprehensiveIncome":0.00,"specialReserve":0.00,"surplusReserve":0.00,"generalRiskReserve":0.00,"undistributedProfit":0.00,"subtotal":0.00,"minorityInterests":0.00,"totalOwnerEquity":0.00},
        {"name":"任意公积金","type":4,"origin":[],"paidInCapital":0.00,"preferedStock":0.00,"perpetualDebt":0.00,"otherEquityInstruments":0.00,"capitalReserve":0.00,"treasuryStock":0.00,"otherComprehensiveIncome":0.00,"specialReserve":0.00,"surplusReserve":0.00,"generalRiskReserve":0.00,"undistributedProfit":0.00,"subtotal":0.00,"minorityInterests":0.00,"totalOwnerEquity":0.00},
        {"name":"#储备基金","type":4,"origin":[],"paidInCapital":0.00,"preferedStock":0.00,"perpetualDebt":0.00,"otherEquityInstruments":0.00,"capitalReserve":0.00,"treasuryStock":0.00,"otherComprehensiveIncome":0.00,"specialReserve":0.00,"surplusReserve":0.00,"generalRiskReserve":0.00,"undistributedProfit":0.00,"subtotal":0.00,"minorityInterests":0.00,"totalOwnerEquity":0.00},
        {"name":"#企业发展基金","type":4,"origin":[],"paidInCapital":0.00,"preferedStock":0.00,"perpetualDebt":0.00,"otherEquityInstruments":0.00,"capitalReserve":0.00,"treasuryStock":0.00,"otherComprehensiveIncome":0.00,"specialReserve":0.00,"surplusReserve":0.00,"generalRiskReserve":0.00,"undistributedProfit":0.00,"subtotal":0.00,"minorityInterests":0.00,"totalOwnerEquity":0.00},
        {"name":"#利润归还投资","type":4,"origin":[],"paidInCapital":0.00,"preferedStock":0.00,"perpetualDebt":0.00,"otherEquityInstruments":0.00,"capitalReserve":0.00,"treasuryStock":0.00,"otherComprehensiveIncome":0.00,"specialReserve":0.00,"surplusReserve":0.00,"generalRiskReserve":0.00,"undistributedProfit":0.00,"subtotal":0.00,"minorityInterests":0.00,"totalOwnerEquity":0.00},
        {"name":"2.提取一般风险准备","type":2,"origin":[],"paidInCapital":0.00,"preferedStock":0.00,"perpetualDebt":0.00,"otherEquityInstruments":0.00,"capitalReserve":0.00,"treasuryStock":0.00,"otherComprehensiveIncome":0.00,"specialReserve":0.00,"surplusReserve":0.00,"generalRiskReserve":0.00,"undistributedProfit":0.00,"subtotal":0.00,"minorityInterests":0.00,"totalOwnerEquity":0.00},
        {"name":"3.对所有者（或股东）的分配","type":2,"origin":[],"paidInCapital":0.00,"preferedStock":0.00,"perpetualDebt":0.00,"otherEquityInstruments":0.00,"capitalReserve":0.00,"treasuryStock":0.00,"otherComprehensiveIncome":0.00,"specialReserve":0.00,"surplusReserve":0.00,"generalRiskReserve":0.00,"undistributedProfit":0.00,"subtotal":0.00,"minorityInterests":0.00,"totalOwnerEquity":0.00},
        {"name":" 4.其他","type":2,"origin":[],"paidInCapital":0.00,"preferedStock":0.00,"perpetualDebt":0.00,"otherEquityInstruments":0.00,"capitalReserve":0.00,"treasuryStock":0.00,"otherComprehensiveIncome":0.00,"specialReserve":0.00,"surplusReserve":0.00,"generalRiskReserve":0.00,"undistributedProfit":0.00,"subtotal":0.00,"minorityInterests":0.00,"totalOwnerEquity":0.00},
        {"name":"（五）所有者权益内部结转","type":1,"origin":[],"paidInCapital":0.00,"preferedStock":0.00,"perpetualDebt":0.00,"otherEquityInstruments":0.00,"capitalReserve":0.00,"treasuryStock":0.00,"otherComprehensiveIncome":0.00,"specialReserve":0.00,"surplusReserve":0.00,"generalRiskReserve":0.00,"undistributedProfit":0.00,"subtotal":0.00,"minorityInterests":0.00,"totalOwnerEquity":0.00},
        {"name":"1.资本公积转增资本（或股本）","type":2,"origin":[],"paidInCapital":0.00,"preferedStock":0.00,"perpetualDebt":0.00,"otherEquityInstruments":0.00,"capitalReserve":0.00,"treasuryStock":0.00,"otherComprehensiveIncome":0.00,"specialReserve":0.00,"surplusReserve":0.00,"generalRiskReserve":0.00,"undistributedProfit":0.00,"subtotal":0.00,"minorityInterests":0.00,"totalOwnerEquity":0.00},
        {"name":"2.盈余公积转增资本（或股本）","type":2,"origin":[],"paidInCapital":0.00,"preferedStock":0.00,"perpetualDebt":0.00,"otherEquityInstruments":0.00,"capitalReserve":0.00,"treasuryStock":0.00,"otherComprehensiveIncome":0.00,"specialReserve":0.00,"surplusReserve":0.00,"generalRiskReserve":0.00,"undistributedProfit":0.00,"subtotal":0.00,"minorityInterests":0.00,"totalOwnerEquity":0.00},
        {"name":"3.盈余公积弥补亏损","type":2,"origin":[],"paidInCapital":0.00,"preferedStock":0.00,"perpetualDebt":0.00,"otherEquityInstruments":0.00,"capitalReserve":0.00,"treasuryStock":0.00,"otherComprehensiveIncome":0.00,"specialReserve":0.00,"surplusReserve":0.00,"generalRiskReserve":0.00,"undistributedProfit":0.00,"subtotal":0.00,"minorityInterests":0.00,"totalOwnerEquity":0.00},
        {"name":"4.设定受益计划变动额结转留存收益","type":2,"origin":[],"paidInCapital":0.00,"preferedStock":0.00,"perpetualDebt":0.00,"otherEquityInstruments":0.00,"capitalReserve":0.00,"treasuryStock":0.00,"otherComprehensiveIncome":0.00,"specialReserve":0.00,"surplusReserve":0.00,"generalRiskReserve":0.00,"undistributedProfit":0.00,"subtotal":0.00,"minorityInterests":0.00,"totalOwnerEquity":0.00},
        {"name":"☆5、其他综合收益结转留存收益","type":2,"origin":[],"paidInCapital":0.00,"preferedStock":0.00,"perpetualDebt":0.00,"otherEquityInstruments":0.00,"capitalReserve":0.00,"treasuryStock":0.00,"otherComprehensiveIncome":0.00,"specialReserve":0.00,"surplusReserve":0.00,"generalRiskReserve":0.00,"undistributedProfit":0.00,"subtotal":0.00,"minorityInterests":0.00,"totalOwnerEquity":0.00},
        {"name":"6.其他","type":2,"origin":[],"paidInCapital":0.00,"preferedStock":0.00,"perpetualDebt":0.00,"otherEquityInstruments":0.00,"capitalReserve":0.00,"treasuryStock":0.00,"otherComprehensiveIncome":0.00,"specialReserve":0.00,"surplusReserve":0.00,"generalRiskReserve":0.00,"undistributedProfit":0.00,"subtotal":0.00,"minorityInterests":0.00,"totalOwnerEquity":0.00},
        {"name":"四、本年年末余额","type":0,"origin":[],"paidInCapital":0.00,"preferedStock":0.00,"perpetualDebt":0.00,"otherEquityInstruments":0.00,"capitalReserve":0.00,"treasuryStock":0.00,"otherComprehensiveIncome":0.00,"specialReserve":0.00,"surplusReserve":0.00,"generalRiskReserve":0.00,"undistributedProfit":0.00,"subtotal":0.00,"minorityInterests":0.00,"totalOwnerEquity":0.00},
]

def addFinancialStatements(name,titles,newRecords,companyName,reportDate,currencyUnit):
    addTableTitle(name)
    addFirstLine(document,companyName,reportDate,currencyUnit)
    document.add_section(start_type=0)
    table = createBorderedTable(document,len(newRecords),5,innerLine="single")
    table.columns[0].width = Cm(7)
    setHeader(table,titles)

    for key in range(1,len(newRecords)):
        table.cell(key, 0).width=Cm(13)
        table.cell(key, 1).width=Cm(2)
        table.cell(key,2).width=Cm(6)
        table.cell(key, 3).width=Cm(6)
        table.cell(key, 4).width=Cm(3)
        setCell(table.cell(key, 0), *getAlignAndText(newRecords[key]["type"],newRecords[key]["name"]), False, "tableSmallCharacter")
        setCell(table.cell(key, 1), key,WD_PARAGRAPH_ALIGNMENT.CENTER, False, "tableSmallCharacter")
        setCell(table.cell(key, 2),newRecords[key]["endDate"],WD_PARAGRAPH_ALIGNMENT.RIGHT, True, "tableSmallCharacter")
        setCell(table.cell(key, 3), newRecords[key]["startDate"],WD_PARAGRAPH_ALIGNMENT.RIGHT, True, "tableSmallCharacter")
        setCell(table.cell(key, 4), "",WD_PARAGRAPH_ALIGNMENT.CENTER, True, "tableSmallCharacter")

    document.add_section(start_type=0)
    add_last_line()
    document.add_section()




def setOwnerHeader(table,period):
        table.cell(0, 0).width = Cm(8)
        table.cell(1, 0).width = Cm(8)
        table.cell(2, 0).width = Cm(8)
        table.cell(3, 0).width = Cm(8)
        table.cell(0, 1).width = Cm(2)
        table.cell(1, 1).width = Cm(2)
        table.cell(2, 1).width = Cm(2)
        table.cell(3, 1).width = Cm(2)

        setCell(table.cell(0,0),"项            目",WD_PARAGRAPH_ALIGNMENT.CENTER,False,"tableSmallerCharacter")
        setCell(table.cell(0,1),"行次",WD_PARAGRAPH_ALIGNMENT.CENTER,False,"tableSmallerCharacter")
        setCell(table.cell(0,2),period,WD_PARAGRAPH_ALIGNMENT.CENTER,False,"tableSmallerCharacter")
        setCell(table.cell(1,2),"归属于母公司所有者权益",WD_PARAGRAPH_ALIGNMENT.CENTER,False,"tableSmallerCharacter")
        setCell(table.cell(1,14),"少数股东权益",WD_PARAGRAPH_ALIGNMENT.CENTER,False,"tableSmallerCharacter")
        setCell(table.cell(1,15),"所有者权益合计",WD_PARAGRAPH_ALIGNMENT.CENTER,False,"tableSmallerCharacter")
        setCell(table.cell(2,2),"实收资本(或股本)",WD_PARAGRAPH_ALIGNMENT.CENTER,False,"tableSmallerCharacter")
        setCell(table.cell(2,3),"其他权益工具",WD_PARAGRAPH_ALIGNMENT.CENTER,False,"tableSmallerCharacter")
        setCell(table.cell(2,6),"资本公积",WD_PARAGRAPH_ALIGNMENT.CENTER,False,"tableSmallerCharacter")
        setCell(table.cell(2,7),"减:库存股",WD_PARAGRAPH_ALIGNMENT.CENTER,False,"tableSmallerCharacter")
        setCell(table.cell(2,8),"其他综合收益",WD_PARAGRAPH_ALIGNMENT.CENTER,False,"tableSmallerCharacter")
        setCell(table.cell(2,9),"专项储备",WD_PARAGRAPH_ALIGNMENT.CENTER,False,"tableSmallerCharacter")
        setCell(table.cell(2,10),"盈余公积",WD_PARAGRAPH_ALIGNMENT.CENTER,False,"tableSmallerCharacter")
        setCell(table.cell(2,11),"△一般风险准备",WD_PARAGRAPH_ALIGNMENT.CENTER,False,"tableSmallerCharacter")
        setCell(table.cell(2,12),"未分配利润",WD_PARAGRAPH_ALIGNMENT.CENTER,False,"tableSmallerCharacter")
        setCell(table.cell(2,13),"小计",WD_PARAGRAPH_ALIGNMENT.CENTER,False,"tableSmallerCharacter")
        setCell(table.cell(3,3),"优先股",WD_PARAGRAPH_ALIGNMENT.CENTER,False,"tableSmallerCharacter")
        setCell(table.cell(3,4),"永续债",WD_PARAGRAPH_ALIGNMENT.CENTER,False,"tableSmallerCharacter")
        setCell(table.cell(3,5),"其他",WD_PARAGRAPH_ALIGNMENT.CENTER,False,"tableSmallerCharacter")
        # 合并单元格
        # 第一列
        for i in range(1,4):
                table.cell(i,0).merge(table.cell(i-1,0))
        # 第二列
        for i in range(1,4):
                table.cell(i,1).merge(table.cell(i-1,1))
        # # 第一行
        for i in range(3,16):
                table.cell(0,i).merge(table.cell(0,i-1))
        # 第二行
        for i in range(13,2,-1):
                table.cell(1,i).merge(table.cell(1,i-1))

        # 第三列
        table.cell(3,2).merge(table.cell(2,2))
        # 第三行
        for i in range(4,6):
                table.cell(2,i).merge(table.cell(2,i-1))
        # 第七列到十三列
        for i in range(6,14):
                table.cell(3,i).merge(table.cell(2,i))
        # 第十四列到十五列
        for i in range(14,16):
                table.cell(2,i).merge(table.cell(1,i))
                table.cell(3,i).merge(table.cell(2,i))

def addOnwerEquity(document,name,newRecords,companyName,reportDate,currencyUnit,period):
    addTableTitle(name)
    addFirstLine(document,companyName,reportDate,currencyUnit)
    document.add_section(start_type=0)
    table = createBorderedTable(document,len(newRecords)+3,16,innerLine="single")
    setOwnerHeader(table,period)

    for key in range(1,len(newRecords)):
        table.cell(key+3, 0).width = Cm(8)
        table.cell(key+3, 1).width = Cm(1)

        setCell(table.cell(key+3, 0), *getAlignAndText(newRecords[key]["type"],newRecords[key]["name"]), False, "tableSmallerCharacter")
        setCell(table.cell(key+3, 1), key, WD_PARAGRAPH_ALIGNMENT.CENTER, False, "tableSmallerCharacter")
        setCell(table.cell(key+3, 2),newRecords[key]["paidInCapital"],WD_PARAGRAPH_ALIGNMENT.RIGHT, True, "tableSmallerCharacter")
        setCell(table.cell(key+3, 3), newRecords[key]["preferedStock"],WD_PARAGRAPH_ALIGNMENT.RIGHT, True, "tableSmallerCharacter")
        setCell(table.cell(key+3, 4), newRecords[key]["perpetualDebt"],WD_PARAGRAPH_ALIGNMENT.RIGHT, True, "tableSmallerCharacter")
        setCell(table.cell(key+3, 5), newRecords[key]["otherEquityInstruments"],WD_PARAGRAPH_ALIGNMENT.RIGHT, True, "tableSmallerCharacter")
        setCell(table.cell(key+3, 6), newRecords[key]["capitalReserve"],WD_PARAGRAPH_ALIGNMENT.RIGHT, True, "tableSmallerCharacter")
        setCell(table.cell(key+3, 7), newRecords[key]["treasuryStock"],WD_PARAGRAPH_ALIGNMENT.RIGHT, True, "tableSmallerCharacter")
        setCell(table.cell(key+3, 8), newRecords[key]["otherComprehensiveIncome"],WD_PARAGRAPH_ALIGNMENT.RIGHT, True, "tableSmallerCharacter")
        setCell(table.cell(key+3, 9), newRecords[key]["specialReserve"],WD_PARAGRAPH_ALIGNMENT.RIGHT, True, "tableSmallerCharacter")
        setCell(table.cell(key+3, 10), newRecords[key]["surplusReserve"],WD_PARAGRAPH_ALIGNMENT.RIGHT, True, "tableSmallerCharacter")
        setCell(table.cell(key+3, 11), newRecords[key]["generalRiskReserve"],WD_PARAGRAPH_ALIGNMENT.RIGHT, True, "tableSmallerCharacter")
        setCell(table.cell(key+3, 12), newRecords[key]["undistributedProfit"],WD_PARAGRAPH_ALIGNMENT.RIGHT, True, "tableSmallerCharacter")
        setCell(table.cell(key+3, 13), newRecords[key]["subtotal"],WD_PARAGRAPH_ALIGNMENT.RIGHT, True, "tableSmallerCharacter")
        setCell(table.cell(key+3, 14), newRecords[key]["minorityInterests"],WD_PARAGRAPH_ALIGNMENT.RIGHT, True, "tableSmallerCharacter")
        setCell(table.cell(key+3, 15), newRecords[key]["totalOwnerEquity"],WD_PARAGRAPH_ALIGNMENT.RIGHT, True, "tableSmallerCharacter")

    document.add_section(start_type=0)
    add_last_line()


# addOnwerEquity("合并所有者权益变动表",ownerRecords,companyName,reportDate,currencyUnit)
def addOwnership(document,ownerRecords,companyName, reportDate, currencyUnit):
    addOnwerEquity(document,"合并所有者权益变动表", ownerRecords, companyName, reportDate, currencyUnit,"本 年 金 额")
    document.add_section()
    addOnwerEquity(document,"合并所有者权益变动表（续）", ownerRecords, companyName, reportDate, currencyUnit,"上 年 金 额")


# 添加合并资产负债表
addFinancialStatements("合并资产负债表",balanceTitles,assetsRecords,companyName,reportDate,currencyUnit)
addFinancialStatements("合并资产负债表(续)",balanceTitles,liabilitiesRecords,companyName,reportDate,currencyUnit)
addFinancialStatements("资产负债表",balanceTitles,assetsRecords,companyName,reportDate,currencyUnit)
addFinancialStatements("资产负债表(续)",balanceTitles,liabilitiesRecords,companyName,reportDate,currencyUnit)
addFinancialStatements("合并利润表",profitTitles,profitRecords,companyName,reportPeriod,currencyUnit)
addFinancialStatements("利润表",profitTitles,profitRecords,companyName,reportPeriod,currencyUnit)
addFinancialStatements("合并现金流量表",profitTitles,cashRecords,companyName,reportPeriod,currencyUnit)
addFinancialStatements("现金流量表",profitTitles,cashRecords,companyName,reportPeriod,currencyUnit)
# 添加所有者权益变动表
addLandscapeContent(document,addOwnership,ownerRecords,companyName,reportDate,currencyUnit)



document.save("fs.docx")

# 上市公司报表
# 已实施新准则
# 未实施新准则
# 首次执行新准则