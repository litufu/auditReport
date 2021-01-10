# -*- coding: UTF-8 -*-

from docx import Document
import pandas as pd

from project.settings import setStyle
from project.utils import addTitle, addParagraph, addTable


# 添加税收
def addTax(document,context,path):
    companyType = context["report_params"]["companyType"]

    if companyType == "国有企业":
        addTitle(document, "六、税项", 1, False)
    else:
        addTitle(document, "五、税项", 1, False)
    addTitle(document, "（一）主要税种及税率", 2, True)
    df = pd.read_excel(path, sheet_name="主要税种及税率")
    dc = df.to_dict("split")
    addTable(document, dc, style=5)
    for content in context["tax"]["policy"]:
        addParagraph(document, content, "paragraph")
    addTitle(document, "（二）税收优惠及批文", 2, True)
    if len(context["tax"]["taxPreference"])==0:
        addParagraph(document, "无。", "paragraph")
    else:
        for content in context["tax"]["taxPreference"]:
            addParagraph(document, content, "paragraph")

def test():
    from project.data import testcontext
    from project.constants import CURRENTPATH

    document = Document()
    setStyle(document)
    addTax(document,testcontext,CURRENTPATH)
    document.save("tax.docx")

if __name__ == '__main__':
    test()
