# -*- coding: UTF-8 -*-

from docx import Document
import pandas as pd

from project.data import tax
from project.settings import setStyle
from project.utils import addTitle, addParagraph,addTable

# 添加税收
def addTax():
    MODELPATH = "D:/auditReport/project/model.xlsx"
    document = Document()
    setStyle(document)
    addTitle(document, "六、税项", 1, False)
    addTitle(document, "（一）主要税种及税率", 2, True)
    df = pd.read_excel(MODELPATH, sheet_name="主要税种及税率")
    dc = df.to_dict("split")
    addTable(document, dc, style=5)
    for content in tax["policy"]:
        addParagraph(document, content, "paragraph")
    addTitle(document, "（二）税收优惠及批文", 2, True)
    for content in tax["taxPreference"]:
        addParagraph(document, content, "paragraph")

    document.save("tax.docx")