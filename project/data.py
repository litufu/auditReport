# -*- coding: UTF-8 -*-

import pandas as pd

# 审计报告
report_params = {
    "type":"单体",#单体、合并
    "companyType": "国有企业",  # 国有企业、上市公司
    "reportNo":"亚会A审字（2020）0580号",
    "companyName":"杭州市城市建设发展集团有限公司",
    "CompanyAbbrName":"城建发展",
    "reportDate":"2019年12月31日",
    "reportPeriod":"2019年度",
    "otherInfo":False,
    "internalControlAudit":False,
    # "keyAuditMatters":[{"title":"(一)应收账款预期信用损失计量","desc":["截至2019年12月31日，如中恒电气公司合并财务报表附注五（三）中恒电气公司应收账款余额1,092,359,178.05元，坏账准备金额133,107,036.74元，账面价值较高。若应收账款不能按期收回或无法收回而发生坏账对财务报表影响较为重大，为此我们将应收账款预期信用损失计量确定为关键审计事项。"],"auditResponse":["针对应收账款预期信用损失计量，我们执行的审计程序主要包括：","（1）了解和评价管理层与应收账款预期信用损失计提相关的关键内部控制的设计和运行有效性。","（2）分析中恒电气公司应收账款预期信用损失计量会计估计的合理性。","（3）复核以前年度应收账款坏账计提后转回或实际发生损失的情况，判断中恒电气公司管理层对历史数据预期的准确性。"]},{"title":"（二）收入确认","desc":["通信电源系统收入、电力电源系统收入和软件开发销售及服务收入是中恒电气公司营业收入的主要组成部分，恰当确认和计量直接关系到年度财务报表的准确性、合理性。为此，我们将收入确认确定为关键审计事项。"],"auditResponse":["针对收入确认，我们执行的审计程序主要包括：","（1） 评价、测试中恒电气公司收入确认相关内部控制的设计和运行有效性。","（2）通过审阅销售合同并与管理层进行访谈，识别与商品所有权上的风险和报酬转移相关的合同条款与条件，评价收入确认时点是否符合企业会计准则的要求。"]}],
    "keyAuditMatters":[],
    "accountFirm":"亚太（集团）会计师事务所（特殊普通合伙）",
    "accountFirmAddr":"中国·北京",
    "issuanceDate":"二〇二〇年三月二十日"
}


# # 财务报表
# fs_params = {
#     "type":"单体",#合并、单体、母公司
# }
# 会计政策
notes_params = {
    "criterion":"新准则",#新准则、老准则、部分新准则
    "tandardssForFinancialInstruments":"新金融工具准则",#新金融工具准则、老金融工具准则、部分新金融工具准则
    "incomeriteria":"新收入准则",#新收入准则、老收入准则、部分新收入工具准则
    "leasingriteria":"新租赁准则",#新租赁准则、老租赁准则、部分新租赁工具准则
    "currencyUnit":"人民币元",
    "companyIntroduce":"杭州市城市建设发展集团有限公司(以下简称“公司”或“本公司”)原系经杭州市机构编制委员会杭编〔1993〕70号文和杭州市城乡建设委员会杭建组发〔1993〕216号文批准组建的实行企业化管理的全民所有制事业单位，注册资本2,000万元。2000年8月，根据杭州市财政局杭财基〔2000〕字511号文，本公司增加注册资本300,000万元，增资后注册资本为302,000万元。公司于2001年12月30日在杭州市工商行政管理局登记注册，取得注册号330100000074750的《企业法人营业执照》。2003年6月，根据中共杭州市委〔2003〕58号文，公司股权划转至杭州市城市建设投资集团有限公司(以下简称杭州城投公司)，成为其下属的全资子公司。2017年9月7日由杭州市市场监督管理局换发统一社会信用代码为91330100470102706G的《营业执照》。法定代表人：胥东；注册地址：杭州市江干区新塘路33-35号12、13层。",
    "companyBussiness":"本公司经营范围：服务：经营出资人授权范围内的国有资产，房地产开发、经营； 服务：市政项目建设、管理、咨询，停车设施的技术开发；批发、零售：建筑材料，金属材料，五金交电，装饰材料，木材，机电设备，百货，工艺美术品，包装材料，羊毛，纸张；货物及技术进出口（法律、行政法规禁止的项目除外，法律、行政法规限制的项目取得许可后方可经营）；其他无需报经审批的一切合法项目；含下属分支机构经营范围。",
    "companyParent":"本公司的母公司为杭州市城市建设投资集团有限公司。",
    "companyIssuedDate":"2020年3月23日",
    "operationPeriod":"营业期限自2004年12月30日至2054年12月29日止。",
    "judgmentStandardOfSignificantSingleAmount":"本公司将同一客户应收款项占年末所属类别应收款项总额10%或以上的应收款项确认为单项金额重大的应收款项。",
    "combinationClassification":[["组合1","除组合2以外的应收款项","账龄分析法"],["组合2","根据业务性质和客户的历史交易情况，认定信用风险不重大的应收款项","不计提坏账"]],
    "badDebtProvisionRatio":[["账龄","应收账款计提比例","其他应收款计提比例"],["1年以内(含1年)","5%","5%"],["1-2年(含2年)","10%","10%"],["1年以内(含1年)","20%","20%"],["2-3年(含3年)","50%","50%"], ["5年以上","100%","100%"]],
    "inventoryClassification":["原材料","库存商品","代建项目","工程施工","开发成本","开发产品","发出商品"],
    "inventoryDeliveryPricing":"加权平均法",#先进先出法/加权平均法/个别认定法等
    "inventoryTaking":"永续盘存制",#永续盘存制/定期盘存制
    "turnoverMaterials":"一次摊销法",#五五摊销法/一次摊销法/分次摊销法
    "measurementModeOfInvestmentRealEstate":"成本模式",#成本模式/公允价值模式
    "fixedAssetsCategory":[["固定资产类别","折旧年限","残值率(%)","年折旧率(%)","折旧方法"],["房屋及建筑物","20-50","0-5","1.90-5.00","年限平均法"],["道路资产","50","0","2.00","年限平均法"]],
    "biologicalAssets":["消耗性生物资产","生产性生物资产"],
    "biologicalAssetsCategory":["牧"],#农林牧渔
    "consumptiveBiologicalAssets":["仔猪","保育猪","育肥猪"],#消耗性生物资产分类
    "productiveBiologicalAssets":["公猪","母猪","种鸭"],#生产性生物资产分类
    "productiveBiologicalAssetsDepreciation": [["资产类别", "折旧方法", "使用寿命(年)", "残值率(%)", "年折旧率(%)"],["公猪", "年限平均法", "3", "30", "85"],["母猪", "年限平均法", "3", "30", "85"]],#生产性生物资产折旧
    "rightToUseAssetsDepreciation": [["资产类别", "折旧方法", "使用寿命(年)", "残值率(%)", "年折旧率(%)"],["房屋及建筑物", "年限平均法", "3", "30", "85"],["专用设备", "年限平均法", "3", "30", "85"]],#使用权资产折旧
    "oldSpecificMethodsOfRevenueRecognition": [
        ["开发产品","已经完工并验收合格，签订了销售合同并履行了合同规定的义务，即开发产品所有权上的主要风险和报酬转移给购货方；此类公司不再保留通常与所有权相联系的继续管理权，也不再对已售出的商品实施有效控制；收入的金额能够可靠地计量；相关的经济利益很可能流入；并且该项目已发生或将发生的成本能够可靠地计量时，确认销售收入的实现。"],
        ["分期收款销售","在开发产品已经完工并验收合格，签订了分期收款销售合同并履行了合同规定的义务，相关的经济利益很可能流入，并且该开发产品成本能够可靠地计量时，按照应收的合同或协议价款的公允价值确定收入金额；应收的合同或协议价款与其公允价值间的差额，在合同或协议期间内采用实际利率法进行摊销，计入当期损益。"],
        ["出售自用房屋","自用房屋所有权上的主要风险和报酬转移给购货方，此类公司不再保留通常与所有权相联系的继续管理权，也不再对已售出的商品实施有效控制，收入的金额能够可靠地计量，相关的经济利益很可能流入，并且该房屋成本能够可靠地计量时，确认销售收入的实现。"],
        ["代建房屋和工程业务","代建房屋和工程签订有不可撤销的建造合同，与代建房屋和工程相关的经济利益能够流入企业，代建房屋和工程的完工进度能够可靠地确定，并且代建房屋和工程有关的成本能够可靠地计量时，采用完工百分比法确认营业收入的实现。"],
        ["其他业务收入确认方法","按相关合同、协议的约定，与交易相关的经济利益很可能流入企业，收入金额能够可靠计量，与收入相关的已发生或将发生成本能够可靠地计量时，确认其他业务收入的实现。"],
    ], #老收入准则确认具体方法
    "newSpecificMethodsOfRevenueRecognition":[
        ["房地产开发业务的收入","本公司商品房销售业务的收入于将物业的控制权转移给客户时确认。基于销售合同条款及适用于合同的法律规定，物业的控制权可在某一时段内或在某一时点转移。仅当本公司在履约过程中所产出的商品具有不可替代用途，且本公司在整个合同期间内有权就累计至今已完成的履约部分收取款项的情况下，按照合同期间已完成履约义务的进度在一段时间内确认收入，已完成履约义务的进度按照为完成履约义务而实际发生的合同成本占合同预计总成本的比例确定。否则，收入于客户获得实物所有权或已完工物业的法定所有权且本公司已获得现时收款权并很可能收回对价时确认。在确认合同交易价格时，若融资成分重大，本公司将根据合同的融资成分来调整合同承诺对价。"],
        ["土地一级开发","本公司一级土地开发业务部分合同按照已完成履约义务的进度在一段时间内确认，已完成履约义务的进度按照为完成履约义务而实际发生的合同成本占合同预计总成本的比例确定；部分合同收入在某一时点确认。"],
        ["物业服务合同","本公司在提供物业服务过程中确认收入。"],
        ["销售商品的收入","本公司销售商品并在客户取得相关商品的控制权时，根据历史经验，按照期望值法确定。折扣金额，按照合同对价扣除预计折扣金额后的净额确认收入。本公司给予客户的信用期与行业惯例一致，不存在重大融资成分。本公司为部分产品提供产品质量保证，并确认相应的预计负债。"],
        ["基础设施建设业务的收入","本公司与客户之间的工程承包合同通常包括房屋建筑建设、基础设施建设等履约义务，由于客户能够控制本公司履约过程中在建的商品，本公司将其作为某一时段内履行的履约义务，根据履约进度在一段时间内确认收入，履约进度不能合理确定的除外。本公司按照投入法确定提供服务的履约进度。对于履约进度不能合理确定时，本公司已经发生的成本预计能够得到补偿的，按照已经发生的成本金额确认收入，直到履约进度能够合理确定为止。合同成本不能收回的，在发生时立即确认为合同费用，不确认合同收入。如果合同总成本很可能超过合同总收入，则形成合同预计损失，计入预计负债，并确认为当期成本。","合同成本包括合同履约成本和合同取得成本。本公司为提供房屋建筑建设、基础设施建设等服务而发生的成本，确认为合同履约成本。本公司在确认收入时，按照已完工的进度将合同履约成本结转计入主营业务成本。本公司将为获取工程承包合同而发生的增量成本，确认为合同取得成本。本公司对于摊销期限不超过一年或者该业务营业周期的合同取得成本，在其发生时计入当期损益；对于摊销期限在一年或一个营业周期以上的合同取得成本，本公司按照相关合同下确认收入相同的基础摊销计入损益。对于初始确认时摊销期限超过一年或者一个营业周期的合同取得成本，以减去相关资产减值准备后的净额，列示为其他非流动资产。"],
        ["建设、运营及移交合同","建设、运营及移交合同项于建设阶段，按照基础设施建设业务的收入所述的会计政策确认基础设施建设服务的收入和成本。基础设施建设服务收入按照收取或有权收取的对价计量，并在确认收入的同时，确认合同资产或无形资产，并对合同安排中的重大融资成分进行会计处理。","合同规定基础设施建成后的一定期间内，本公司可以无条件地自合同授予方收取确定金额的货币资金或其他金融资产的，于项目建造完成时，将合同资产转入金融资产核算；","合同规定本公司在有关基础设施建成后，从事经营的一定期间内有权利向获取服务的对象收取费用，但收费金额不确定的，该权利不构成一项无条件收取现金的权利，本公司在确认收入的同时确认无形资产。并在该项目竣工验收之日起至运营期及其延展期届满或特许经营权终止之日的期间采用年限平均法摊销。","于运营阶段，当提供劳务时，确认相应的收入；发生的日常维护或修理费用，确认为当期费用。"],
        ["建设和移交合同","对于本公司提供基础设施建设服务的，于建设阶段，按照基础设施建设业务的收入所述的会计政策确认相关基础设施建设服务收入和成本，基础设施建设服务收入按照收取或有权收取的对价计量，在确认收入的同时确认合同资产，并对合同安排中的重大融资成分进行会计处理。待拥有无条件收取对价权利时，转入“长期应收款”，待收到业主支付的款项后，进行冲减。"],
    ],
    "constructionContractCompletionProgress":"累计实际发生的合同成本占合同预计总成本的比例",#建造合同完工进度 累计实际发生的合同成本占合同预计总成本的比例/已经完成的合同工作量占合同预计总工作量的比例/实际测定的完工进度
    "shareBasedPayment":False,#True/False



}

standardChange = {
    "implementationOfNewStandardsInThisPeriod":["新金融工具准则","新收入准则","新租赁准则"],#本期首次执行新准则，"新金融工具准则",
    "newFinancialInstrumentsChange":[
        "公司于2021年1月1日及以后将持有的部分非交易性股权投资指定为以公允价值计量且其变动计入其他综合收益的金融资产，列报为其他权益工具投资。",
        "公司持有的某些理财产品、信托产品、股权收益权及资产管理计划等，其收益取决于标的资产的收益率，原分类为可供出售金融资产。由于其合同现金流量不仅仅为对本金和以未偿付本金为基础的利息的支付，公司在2021年1月1日及以后将其重分类为以公允价值计量且其变动计入当期损益的金融资产，列报为交易性金融资产。",
        "公司持有的部分可供出售债务工具，其在特定日期产生的现金流量仅为对本金和以未偿付本金金额为基础的利息的支付，且公司管理该金融资产的业务模式是既以收取合同现金流量为目标又以出售该金融资产为目标，本公司在2021年1月1日及以后将其从可供出售金融资产重分类至其他债权投资。",
        "公司在日常资金管理中将部分银行承兑汇票背书或贴现，既以收取合同现金流量又以出售金融资产为目标，因此，公司在2021年1月1日及以后将该等应收票据重分类为以公允价值计量且其变动计入其他综合收益金融资产类别，列报为应收款项融资。",
        "公司在日常资金管理中将部分特定客户的应收款项通过无追索权保理进行出售，针对该部分特定客户的应收款项，既以收取合同现金流量又以出售金融资产为目标，因此，公司在2021年1月1日及以后将该等特定客户的应收款项重分类为以公允价值计量且其变动计入其他综合收益金融资产类别，列报为其他债权投资或应收款项融资。",
        "于2020年12月31日，本公司持有的有息拆借及委托贷款,由于管理该组债权投资的业务模式是以收取合同现金流为目标，且其合同现金流量特征与基本借贷安排相一致，故于2021年1月1日，本集团及本公司将该债权投资从应收款项重分类至以摊余成本计量的金融资产，列示为债权投资、其他流动资产和一年内到期的非流动资产。相应地，本公司按照预期信用损失计量损失准备转出至期初留存收益。",
        "于2020年12月31日，本公司持有非上市信托产品由于其合同现金流量特征不符合基本借贷安排，故于2021年1月1日，本公司将该等信托产品从可供出售金融资产重分类为以公允价值计量且其变动计入当期损益的金融资产，列示为交易性金融资产及其他非流动金融资产。相应地，本集团将累计计入其他综合收益的金额转至期初留存收益。"
        "于2020年12月31日，本公司持有上市权益工具投资由于该等权益工具投资不符合本金加利息的合同现金流量特征，故于2021年1月1日，本公司将此等权益投资从可供出售金融资产重分类为以公允价值计量且其变动计入当期损益的金融资产，列示为交易性金融资产及其他非流动金融资产。相应地，本公司将累计计入其他综合收益的金额转出至期初留存收益。"
        "于2020年12月31日，本公司持有的非上市信托产品由于该等信托产品投资的业务模式为既以收取合同现金流量为目标又以出售为目标，且其合同现金流量特征与基本借贷安排相一致，故于2021年1月1日，本公司将其重分类至其他债权投资及一年内到期的非流动资产。",
        "于2020年12月31日，本公司持有的以成本计量的非上市股权投资于2021年1月1日，出于战略投资的考虑，本公司选择将该等股权投资指定为以公允价值计量且其变动计入其他综合收益的金融资产，列示为其他权益工具投资。相应地，本公司将公允价值与原账面价值的差额调整期初其他综合收益；将累计计提的减值准备从期初留存收益转入其他综合收益。",
    ],
    "newIncomeCriteria":[
        "公司的XX业务原按照完工百分比法分期确认收入，执行新收入准则后，由于不满足在一段时间内确认收入的条件，变更为在商品控制权转让给客户之时一次性确认收入。",
        "公司将因转让商品而预先收取客户的合同对价从“预收账款”项目变更为“合同负债”项目列报。",
        "公司对于客户奖励积分的分摊方法由剩余价值法变更为按照提供商品或服务以及奖励积分单独售价的相对比例进行分摊。",
        "公司的一些应收款项不满足无条件（即：仅取决于时间流逝）向客户收取对价的条件，本公司将其重分类列报为合同资产（或其他非流动资产）；本公司将未到收款期的应收质保金重分类为合同资产（或其他非流动资产）列报。"
        "公司支付给客户（或消费者）的XX费用，原计入销售费用，在新收入准则下作为应付客户对价，冲减营业收入。"
        "公司向客户提供的质量保证服务，原作为预计负债核算，在新收入准则下因向客户提供了所销售商品符合既定标准之外的额外服务，被识别为单项履约义务，在相关服务履行时确认收入。",
        "公司将与基础设施建设、钢结构产品制造与安装业务及提供劳务相关、不满足无条件收款权的已完工未结算、长期应收款计入合同资产和其他非流动资产",
        "将与基础设施建设、钢结构产品制造与安装业务相关的已结算未完工，提供劳务及与销售商品相关的预收款项重分类至合同负债",
        "将与基础设施建设、钢结构产品制造与安装业务及提供劳务相关的合同预计损失准备重分类至预计负债",
    ],
    "newLeaseCriteria":[
        "公司承租XX公司的XXX资产，租赁期为XXX，原作为经营租赁处理，根据新租赁准则，于2019年1月1日确认使用权资产XXX元，租赁负债XXX元。",
        "公司承租XX公司的XXX资产，租赁期为XXX，原作为融资租赁处理，根据新租赁准则，于2019年1月1日将原在固定资产中列报的“融资租入固定资产” XXX元重分类至使用权资产列报，将在长期应付款中列报的“应付融资租赁款”XXX元重分类至租赁负债列报。",
    ],
    "incrementalBorrowingRate":"4.61",
}

tax = {
    "policy":["本公司发生增值税应税销售行为或者进口货物，于2019年1～3月期间的适用税率为16%/10%，根据《财政部、国家税务总局、海关总署关于深化增值税改革有关政策的公告》（财政部、国家税务总局、海关总署公告[2019]39号）规定，自2019年4月1日起，适用税率调整为13%/9%。同时，本公司/XX子公司作为生产性服务业纳税人，自2019年4月1日至2021年12月31日，按照当期可抵扣进项税额加计10%抵减应纳税额；本公司/XX子公司作为生活性服务业纳税人，自2019年4月1日至2019年9月30日按照当期可抵扣进项税额加计10%抵减应纳税额，自2019年10月1日至2021年12月31日按照当期可抵扣进项税额加计15%抵减应纳税额。"],
    "taxPreference":["根据【国家机关】【批准文件文号】，本公司/XX子公司自【20XX】年起至【20XX】年减半按照XX%税率征收企业所得税。","【根据[国家机关][文件编号]号文，本集团下属XX公司，按《XX》的有关规定享受企业所得税优惠政策，即自首个获利年度起，XX年至XX年免缴企业所得税，自XX年至XX年减半缴纳企业所得税，本年度适用税率为XX%。】","【根据[（主管税务机关）（批准文件编号）]号文，本集团下属XX公司获准缓缴[企业所得税]。】"]
}

combine = {
    "statementOnInconsistencyOfAccountingPeriodBetweenSubsidiaryCompanyAndParentCompany":"不适用",#（子公司与母公司会计期间不一致的说明
    "majorRestrictionsOnSubsidiariesUseOfEnterpriseGroupAssetsAndSettlementOfEnterpriseGroupDebts":"不适用",#子企业使用企业集团资产和清偿企业集团债务的重大限制
    "structuredSubject":"不适用",#（十三）	纳入合并财务报表范围的结构化主体的相关信息
    "changesInShareOfOwnerEquityOfParentCompanyInSubsidiaryEnterprises":"不适用",#母公司在子企业的所有者权益份额发生变化的情况
    "theAbilityOfSubsidiaryToTransferFundsToItsParentCompanyIsStrictlyRestricted":"不适用",#子公司向母公司转移资金的能力受到严格限制的情况
}

noteAppend = {
    "newFinancialInstruments":1,#0:老准则，1：新准则
    "newIncomeCriteria":0,#0:老准则，1：新准则
    "newLeaseCriteria":0,#0:老准则，1：新准则
    "purchaseSubsidiaries":False,#购买子公司
    "disposalSubsidiaries":False,#处置子公司
    "foreignCurrencyMonetaryItems":False, #外币货币性项目
    "limitAsset":True,  #受限资产
    "shareBasedPayment":True,  #股份支付
    "SegmentInformation":True,  #分部信息
    "relationTransactionPrice":"本公司销售给关联方的产品、向关联方提供劳务、从关联方购买商品、接受关联方劳务价格参考市场价格经双方协商后确定。"
}

# 获取值
def getValue(name,data):
    for item in data:
        if item[0]==name:
            value = item[1]
            if isinstance(value,pd.Timestamp):
                return "{}年{}月{}日".format(value.year,value.month,value.day)
            elif isinstance(value,float):
                return "{}".format(round(value, 2))
            else:
                value = str.strip(value)
                if value=="无":
                    return False
                elif value=="有":
                    return True
                elif value=="是":
                    return True
                elif value=="否":
                    return False
                elif value.startswith("[") and value.endswith("]"):
                    return eval(value)
                else:
                    return value

def initData(path):
    df = pd.read_excel(path, sheet_name="基础信息",header=None)
    dc = df.to_dict("split")
    data = dc["data"]

    report_params["type"]=getValue("报告类型",data)
    report_params["companyType"]=getValue("公司类型",data)
    report_params["reportNo"]=getValue("审计报告编号",data)
    report_params["companyName"]=getValue("单位名称",data)
    report_params["CompanyAbbrName"]=getValue("公司简称",data)
    report_params["reportDate"]=getValue("报表日期",data)
    report_params["reportPeriod"]=getValue("报表期间",data)
    report_params["otherInfo"]=getValue("报告中其他信息",data)
    report_params["internalControlAudit"]=getValue("内部控制审计",data)
    report_params["keyAuditMatters"]=getValue("关键审计事项",data)
    report_params["accountFirm"]=getValue("会计师事务所",data)
    report_params["accountFirmAddr"]=getValue("会计事务所地址",data)
    report_params["issuanceDate"]=getValue("审计报告日期",data)

    # fs_params["type"] = getValue("本报表类型",data)

    notes_params["tandardssForFinancialInstruments"]=getValue("金融工具准则",data)
    notes_params["incomeriteria"]=getValue("收入准则",data)
    notes_params["leasingriteria"]=getValue("租赁准则",data)
    notes_params["currencyUnit"]=getValue("记账本位币",data)
    notes_params["companyIntroduce"]=getValue("公司简介",data)
    notes_params["companyBussiness"]=getValue("经营范围",data)
    notes_params["companyParent"]=getValue("母公司",data)
    notes_params["companyIssuedDate"]=getValue("董事会批准日期",data)
    notes_params["operationPeriod"]=getValue("营业期限",data)
    notes_params["judgmentStandardOfSignificantSingleAmount"]=getValue("单项金额重大认定标准",data)
    notes_params["combinationClassification"]=getValue("组合分类依据",data)
    notes_params["badDebtProvisionRatio"]=getValue("坏账计提比例",data)
    notes_params["inventoryClassification"]=getValue("存货种类",data)
    notes_params["inventoryDeliveryPricing"]=getValue("存货发出计价方法",data)
    notes_params["inventoryTaking"]=getValue("盘点方法",data)
    notes_params["turnoverMaterials"]=getValue("低值易耗品摊销方法",data)
    notes_params["measurementModeOfInvestmentRealEstate"]=getValue("投资性房地产核算模式",data)
    notes_params["fixedAssetsCategory"]=getValue("固定资产折旧政策",data)
    notes_params["biologicalAssets"]=getValue("生物资产种类",data)
    notes_params["biologicalAssetsCategory"]=getValue("生物资产分类",data)
    notes_params["consumptiveBiologicalAssets"]=getValue("消耗性生物资产分类",data)
    notes_params["productiveBiologicalAssets"]=getValue("生产性生物资产分类",data)
    notes_params["productiveBiologicalAssetsDepreciation"]=getValue("生产性生物资产折旧方法",data)
    notes_params["rightToUseAssetsDepreciation"]=getValue("使用权资产折旧方法",data)
    notes_params["oldSpecificMethodsOfRevenueRecognition"]=getValue("老准则收入确认具体方法",data)
    notes_params["newSpecificMethodsOfRevenueRecognition"]=getValue("新准则收入确认具体方法",data)
    notes_params["constructionContractCompletionProgress"]=getValue("建造合同完工进度",data)
    notes_params["shareBasedPayment"]=getValue("股份支付",data)

    standardChange["implementationOfNewStandardsInThisPeriod"]=getValue("本期执行的新准则",data)
    standardChange["newFinancialInstrumentsChange"]=getValue("新金融工具准则执行变动",data)
    standardChange["newIncomeCriteria"]=getValue("新收入准则执行变动",data)
    standardChange["newLeaseCriteria"]=getValue("新租赁准则执行变动",data)
    standardChange["incrementalBorrowingRate"]=getValue("增量借款利率",data)

    tax["policy"]=getValue("税收政策",data)
    tax["taxPreference"]=getValue("税收优惠",data)

    contrastSubjectToNum={
        "新金融工具准则":1,
        "老金融工具准则":0,
        "部分新金融工具准则":1,
        "部分新收入准则":1,
        "新收入准则":1,
        "老收入准则":0,
        "老租赁准则":0,
        "新租赁准则":1,
        "部分新租赁准则":1,
    }
    noteAppend["newFinancialInstruments"]=contrastSubjectToNum[getValue("金融工具准则",data)]
    noteAppend["newIncomeCriteria"]=contrastSubjectToNum[getValue("收入准则",data)]
    noteAppend["newLeaseCriteria"]=contrastSubjectToNum[getValue("租赁准则",data)]
    noteAppend["purchaseSubsidiaries"]=getValue("购买子公司",data)
    noteAppend["disposalSubsidiaries"]=getValue("处置子公司",data)
    noteAppend["foreignCurrencyMonetaryItems"]=getValue("外币货币性项目",data)
    noteAppend["limitAsset"]=getValue("受限资产",data)
    noteAppend["shareBasedPayment"]=getValue("股份支付",data)
    noteAppend["SegmentInformation"]=getValue("分部信息",data)
    noteAppend["relationTransactionPrice"]=getValue("关联方交易定价",data)

    context = {
        "report_params": report_params,
        # "fs_params": fs_params,
        "notes_params": notes_params,
        "standardChange": standardChange,
        "tax": tax,
        "combine": combine,
        "noteAppend": noteAppend,
    }
    return context


testcontext = {
        "report_params": report_params,
        # "fs_params": fs_params,
        "notes_params": notes_params,
        "standardChange": standardChange,
        "tax": tax,
        "combine": combine,
        "noteAppend": noteAppend,
    }


def test():
    from project.constants import CURRENTPATH
    initData(CURRENTPATH)

if __name__ == '__main__':
    test()