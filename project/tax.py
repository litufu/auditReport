# -*- coding: UTF-8 -*-

from docx import Document
import pandas as pd

from project.data import tax, report_params
from project.settings import setStyle
from project.utils import addTitle, addParagraph, addTable


# 添加税收
def addTax(document):
    companyType = report_params["companyType"]
    MODELPATH = "D:/auditReport/project/model.xlsx"

    if companyType == "国有企业":
        addTitle(document, "六、税项", 1, False)
    else:
        addTitle(document, "五、税项", 1, False)
    addTitle(document, "（一）主要税种及税率", 2, True)
    df = pd.read_excel(MODELPATH, sheet_name="主要税种及税率")
    dc = df.to_dict("split")
    addTable(document, dc, style=5)
    for content in tax["policy"]:
        addParagraph(document, content, "paragraph")
    addTitle(document, "（二）税收优惠及批文", 2, True)
    for content in tax["taxPreference"]:
        addParagraph(document, content, "paragraph")

def test():
    document = Document()
    setStyle(document)
    addTax(document)
    document.save("tax.docx")

if __name__ == '__main__':
    test()
