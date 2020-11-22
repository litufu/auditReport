from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt,Cm
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.section import WD_ORIENTATION,WD_SECTION_START

# 判断字符串是否为数字
def is_number(s):
    try:
        float(s)
        return True
    except Exception:
        pass
    return False

def create_element(name):
    return OxmlElement(name)


def create_attribute(element, name, value):
    element.set(qn(name), value)


def add_page_number(paragraph):
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    page_run = paragraph.add_run(style="zero")
    t1 = create_element('w:t')
    create_attribute(t1, 'xml:space', 'preserve')
    t1.text = '第 '
    page_run._r.append(t1)

    page_num_run = paragraph.add_run()

    fldChar1 = create_element('w:fldChar')
    create_attribute(fldChar1, 'w:fldCharType', 'begin')

    instrText = create_element('w:instrText')
    create_attribute(instrText, 'xml:space', 'preserve')
    instrText.text = "PAGE"

    fldChar2 = create_element('w:fldChar')
    create_attribute(fldChar2, 'w:fldCharType', 'end')

    page_num_run._r.append(fldChar1)
    page_num_run._r.append(instrText)
    page_num_run._r.append(fldChar2)

    of_run = paragraph.add_run(style="zero")
    t2 = create_element('w:t')
    create_attribute(t2, 'xml:space', 'preserve')
    t2.text = ' 页 共 '
    of_run._r.append(t2)

    fldChar3 = create_element('w:fldChar')
    create_attribute(fldChar3, 'w:fldCharType', 'begin')

    instrText2 = create_element('w:instrText')
    create_attribute(instrText2, 'xml:space', 'preserve')
    instrText2.text = "NUMPAGES"

    fldChar4 = create_element('w:fldChar')
    create_attribute(fldChar4, 'w:fldCharType', 'end')

    num_pages_run = paragraph.add_run()
    num_pages_run._r.append(fldChar3)
    num_pages_run._r.append(instrText2)
    num_pages_run._r.append(fldChar4)

    end_run = paragraph.add_run(style="zero")
    t2 = create_element('w:t')
    create_attribute(t2, 'xml:space', 'preserve')
    t2.text = ' 页'
    end_run._r.append(t2)


# 设置单元格边框
def set_cell_border(cell, **kwargs):
    """
    Set cell`s border
    Usage:
    set_cell_border(
        cell,
        top={"sz": 12, "val": "single", "color": "#FF0000", "space": "0"},
        bottom={"sz": 12, "color": "#00FF00", "val": "single"},
        left={"sz": 24, "val": "dashed", "shadow": "true"},
        right={"sz": 12, "val": "dashed"},
    )
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    # check for tag existnace, if none found, then create one
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)

    # list over all available tags
    for edge in ('left', 'top', 'right', 'bottom', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)

            # check for tag existnace, if none found, then create one
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)

            # looks like order of attributes is important
            for key in ["sz", "val", "color", "space", "shadow"]:
                if key in edge_data:
                    element.set(qn('w:{}'.format(key)), str(edge_data[key]))

_MAPPING = ('零','一', '二', '三', '四', '五', '六', '七', '八', '九', '十', '十一', '十二', '十三', '十四', '十五', '十六', '十七','十八', '十九')
_P0 = ('', '十', '百', '千',)
_S4 = 10 ** 4
def to_chinese(num):
    if num < 20:
        return _MAPPING[num]
    else:
        lst = []
        while num >= 10:
            lst.append(num % 10)
            num = num / 10
        lst.append(num)
        # c = len(lst)  # 位数
        result = ''

        for idx, val in enumerate(lst):
            val = int(val)
            if val != 0:
                result += _P0[idx] + _MAPPING[val]
        return result[::-1]


# 添加标题头
def addTitle(document, title, level, indent):
    basic_title = document.add_heading("", level=level)
    if level >= 3:
        basic_title.add_run(title, style="first")
    else:
        basic_title.add_run(title, style="first").bold = True
    basic_title.style.paragraph_format.line_spacing = 1.5
    basic_title.style.paragraph_format.space_before = Pt(0)
    if indent:
        basic_title.style.paragraph_format.first_line_indent = Pt(24)


# 添加段落
def addParagraph(document, content, paragraphStyle):
    paragraph = document.add_paragraph(style=paragraphStyle)
    paragraph.add_run(content, style="zero")

 # 添加表格边框
def createBorderedTable(document,rowLength,columnLength,innerLine="dotted"):
    table = document.add_table(rowLength, columnLength)
    for cell in table.rows[0].cells:
        set_cell_border(cell, top={"sz": 12, "val": "single", "space": "0"})
    for cell in table.rows[-1].cells:
        set_cell_border(cell, bottom={"sz": 12, "val": "single", "space": "0"})
    for row in table.rows[0:len(table.rows) - 1]:
        for cell in row.cells:
            set_cell_border(cell, bottom={"sz": 6, "val": innerLine, "space": "0"})
    for key, column in enumerate(table.columns):
        if key == len(table.columns) - 1:
            continue
        for cell in column.cells:
            set_cell_border(cell, right={"sz": 6, "val": innerLine, "space": "0"})
    # 设置表格格式
    table.autofit = True
    return table

# 设置单元格格式

def setCell(cell,cellText,alignment,toFloat=True,style="tableCharacter"):
    if is_number(cellText) and toFloat:
        cellText =  '{:,.2f}'.format(float(cellText))
    cellString = str(cellText)
    if cellString == "nan":
        cellString = ""

    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    p = cell.paragraphs[0]
    paragraph_format = p.paragraph_format
    # paragraph_format.line_spacing = 1.5
    paragraph_format.space_after = Pt(1)
    paragraph_format.space_before = Pt(1)
    paragraph_format.alignment = alignment
    p.add_run(str(cellString), style=style)


def setNotToFloatCell(title,cell,j,row,alignment):
    if title[j] in ["级次", "序号"]:
        setCell(cell, row[j], alignment, toFloat=False)
    else:
        setCell(cell, row[j], alignment, toFloat=True)

# 向表格中添加数据
'''
style:
1:单元格全部左对齐
2：标题居中，第一列除最后单元格居中外其余全部左对齐，其他列右对齐
3：标题居中，第一列左对齐，其余右对齐
4:全部居中
5：标题居中，其余左对齐
6：数字靠右，文字靠左，标题居中，第一列靠左
'''
# 仅适用于单行标题
def addTable(document, table,style=1):
    # 获取标题和单元格
    title = table["columns"]
    cells = table["data"]
    if len(cells)==0:
        addParagraph(document,"不适用","paragraph")
        return
    columnLength = len(cells[0])
    # 没有标题
    if len(title) == 0:
        data = cells
        rowLength = len(cells)
        table = createBorderedTable(document, rowLength, columnLength)
        if style == 1:
            for i, row in enumerate(data):
                for j, cell in enumerate(table.rows[i].cells):
                    setCell(cell, row[j], WD_PARAGRAPH_ALIGNMENT.LEFT)
        elif style == 2:
            for i, row in enumerate(data):
                for j, cell in enumerate(table.rows[i].cells):
                    if j==0 and i!=rowLength-1:
                        setCell(cell, row[j], WD_PARAGRAPH_ALIGNMENT.LEFT)
                    elif j==0 and i == rowLength-1:
                        setCell(cell, row[j], WD_PARAGRAPH_ALIGNMENT.CENTER)
                    else:
                        setCell(cell, row[j], WD_PARAGRAPH_ALIGNMENT.RIGHT)
        elif style == 3:
            for i, row in enumerate(data):
                for j, cell in enumerate(table.rows[i].cells):
                    if j==0 :
                        setCell(cell, row[j], WD_PARAGRAPH_ALIGNMENT.LEFT)
                    else:
                        setCell(cell, row[j], WD_PARAGRAPH_ALIGNMENT.RIGHT)
        elif style==4:
            for i, row in enumerate(data):
                for j, cell in enumerate(table.rows[i].cells):
                    setCell(cell, row[j], WD_PARAGRAPH_ALIGNMENT.CENTER)
        elif style == 5:
            for i, row in enumerate(data):
                for j, cell in enumerate(table.rows[i].cells):
                    setCell(cell, row[j], WD_PARAGRAPH_ALIGNMENT.LEFT)
        elif style==6:
            for i, row in enumerate(data):
                for j, cell in enumerate(table.rows[i].cells):
                    if j==0 :
                        setCell(cell, row[j], WD_PARAGRAPH_ALIGNMENT.LEFT)
                    else:
                        if is_number(row[j]):
                            setCell(cell, row[j], WD_PARAGRAPH_ALIGNMENT.RIGHT)
                        else:
                            setCell(cell, row[j], WD_PARAGRAPH_ALIGNMENT.LEFT)
    else:
        data = [title, *cells]
        rowLength = len(cells) + 1
        table = createBorderedTable(document, rowLength, columnLength)
        if style == 1:
            for i, row in enumerate(data):
                for j, cell in enumerate(table.rows[i].cells):
                    setNotToFloatCell(title, cell, j, row, WD_PARAGRAPH_ALIGNMENT.LEFT)
        elif style == 2:
            for i, row in enumerate(data):
                for j, cell in enumerate(table.rows[i].cells):
                    if i == 0:
                        setNotToFloatCell(title, cell, j, row, WD_PARAGRAPH_ALIGNMENT.CENTER)
                    else:
                        if j==0 and i!=rowLength-1:
                            setNotToFloatCell(title, cell, j, row, WD_PARAGRAPH_ALIGNMENT.LEFT)
                        elif j==0 and i == rowLength-1:
                            setNotToFloatCell(title, cell, j, row, WD_PARAGRAPH_ALIGNMENT.CENTER)
                        else:
                            setNotToFloatCell(title, cell, j, row, WD_PARAGRAPH_ALIGNMENT.RIGHT)
        elif style ==3:
            for i, row in enumerate(data):
                for j, cell in enumerate(table.rows[i].cells):
                    if i == 0:
                        setNotToFloatCell(title, cell, j, row, WD_PARAGRAPH_ALIGNMENT.CENTER)
                    else:
                        if j==0 :
                            setNotToFloatCell(title, cell, j, row, WD_PARAGRAPH_ALIGNMENT.LEFT)
                        else:
                            setNotToFloatCell(title, cell, j, row, WD_PARAGRAPH_ALIGNMENT.RIGHT)
        elif style==4:
            for i, row in enumerate(data):
                for j, cell in enumerate(table.rows[i].cells):
                    setNotToFloatCell(title, cell, j, row, WD_PARAGRAPH_ALIGNMENT.CENTER)
        elif style ==5:
            for i, row in enumerate(data):
                for j, cell in enumerate(table.rows[i].cells):
                    if i == 0:
                        setNotToFloatCell(title, cell, j, row, WD_PARAGRAPH_ALIGNMENT.CENTER)
                    else:
                        setNotToFloatCell(title, cell, j, row, WD_PARAGRAPH_ALIGNMENT.LEFT)
        elif style == 6:
            for i, row in enumerate(data):
                for j, cell in enumerate(table.rows[i].cells):
                    if i == 0:
                        setNotToFloatCell(title, cell, j, row, WD_PARAGRAPH_ALIGNMENT.CENTER)
                    else:
                        if j == 0:
                            setNotToFloatCell(title, cell, j, row, WD_PARAGRAPH_ALIGNMENT.LEFT)
                        else:
                            if is_number(row[j]):
                                setNotToFloatCell(title, cell, j, row, WD_PARAGRAPH_ALIGNMENT.RIGHT)
                            else:
                                setNotToFloatCell(title, cell, j, row, WD_PARAGRAPH_ALIGNMENT.LEFT)


# 添加众向合并表格，自动合并下一个为空的单元格
def addCombineTableContent(table,dc,titleLength):
    for key in range(len(dc["index"])):
        row = key + titleLength
        for j in range(len(dc["columns"])):
            cell = table.cell(row, j)
            if str(dc["data"][key][j]) == "nan":
                cell.merge(table.cell(row - 1, j))
            else:
                setCell(cell,str(dc["data"][key][j]),WD_PARAGRAPH_ALIGNMENT.LEFT)

# 添加众向合并表格，自动合并下一个为空的单元格
def addContentToCombineTitle(document,dc,table,titleLength,style=1):
    # 获取标题和单元格
    cells = dc["data"]
    if len(cells) == 0:
        addParagraph(document, "不适用", "paragraph")
        return
    # 没有标题
    data = cells
    rowLength = len(cells)+titleLength
    if style == 1:
        for i, row in enumerate(data):
            i = i+titleLength
            for j, cell in enumerate(table.rows[i].cells):
                setCell(cell, row[j], WD_PARAGRAPH_ALIGNMENT.LEFT)
    elif style == 2:
        for i, row in enumerate(data):
            i = i + titleLength
            for j, cell in enumerate(table.rows[i].cells):
                if j == 0 and i != rowLength - 1:
                    setCell(cell, row[j], WD_PARAGRAPH_ALIGNMENT.LEFT)
                elif j == 0 and i == rowLength - 1:
                    setCell(cell, row[j], WD_PARAGRAPH_ALIGNMENT.CENTER)
                else:
                    setCell(cell, row[j], WD_PARAGRAPH_ALIGNMENT.RIGHT)
    elif style == 3:
        for i, row in enumerate(data):
            i = i + titleLength
            for j, cell in enumerate(table.rows[i].cells):
                if j == 0:
                    setCell(cell, row[j], WD_PARAGRAPH_ALIGNMENT.LEFT)
                else:
                    setCell(cell, row[j], WD_PARAGRAPH_ALIGNMENT.RIGHT)
    elif style == 4:
        for i, row in enumerate(data):
            i = i + titleLength
            for j, cell in enumerate(table.rows[i].cells):
                setCell(cell, row[j], WD_PARAGRAPH_ALIGNMENT.CENTER)
    elif style == 5:
        for i, row in enumerate(data):
            for j, cell in enumerate(table.rows[i].cells):
                setCell(cell, row[j], WD_PARAGRAPH_ALIGNMENT.LEFT)
    elif style == 6:
        for i, row in enumerate(data):
            i = i + titleLength
            for j, cell in enumerate(table.rows[i].cells):
                if j == 0:
                    setCell(cell, row[j], WD_PARAGRAPH_ALIGNMENT.LEFT)
                else:
                    if is_number(row[j]):
                        setCell(cell, row[j], WD_PARAGRAPH_ALIGNMENT.RIGHT)
                    else:
                        setCell(cell, row[j], WD_PARAGRAPH_ALIGNMENT.LEFT)

# 检查标题左面是否全部为空
def checkLeftSpace(row,j):
    for i in range(j):
        if row[i]!="nan":
            return False
    return True

# 添加合并标题
def addCombineTableTitle(table,titles):
    for i,row in enumerate(titles):
        for j,cellText in enumerate(row):
            cell = table.cell(i,j)
            if cellText=="nan" :
                if j>=1:
                    if checkLeftSpace(row,j):
                        cell.merge(table.cell(i - 1, j))
                    else:
                        cell.merge(table.cell(i,j-1))
                else:
                    cell.merge(table.cell(i-1,j))
            else:
                setCell(cell,cellText,WD_PARAGRAPH_ALIGNMENT.CENTER)

# 添加横向内容
'''
type:合并、单体
'''
def addLandscapeContent(document,func,*args):
    # 设置横向
    section = document.add_section(start_type=WD_SECTION_START.CONTINUOUS)
    section.orientation = WD_ORIENTATION.LANDSCAPE
    page_h, page_w = section.page_width, section.page_height
    section.page_width = page_w
    section.page_height = page_h

    # 添加内容
    func(document,*args)

    # 重新设置为众向
    section = document.add_section(start_type=WD_SECTION_START.CONTINUOUS)
    section.orientation = WD_ORIENTATION.PORTRAIT
    page_h, page_w = section.page_width, section.page_height
    section.page_width = page_w
    section.page_height = page_h