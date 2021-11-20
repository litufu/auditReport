from docx.oxml.ns import qn
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt,Cm
from docx.shared import RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT



def setStyle(document):
    # 设置A4纸
    section = document.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)

    style_song = document.styles.add_style('Song', WD_STYLE_TYPE.CHARACTER)  # 设置Song字样式
    style_song.font.name = '宋体'
    document.styles['Song']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')  # 将段落中的所有字体

    # 正文标题
    style_title = document.styles.add_style('title', WD_STYLE_TYPE.CHARACTER)  # 设置一级字样式
    style_title.font.name = '宋体'
    style_title.font.bold = True
    style_title.font.size = Pt(15)
    style_title.font.color.rgb = RGBColor(0x0, 0x0, 0x0)
    document.styles['title']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')  # 将段落中的所有字体

    # 一级标题：宋体 加粗 12号字体
    style_first = document.styles.add_style('first', WD_STYLE_TYPE.CHARACTER)  # 设置一级字样式
    style_first.font.name = '宋体'
    style_first.font.bold = True
    style_first.font.size = Pt(12)
    style_first.font.color.rgb = RGBColor(0x0, 0x0, 0x0)
    document.styles['first']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')  # 将段落中的所有字体

    # 设置表格字体
    tableCharacter = document.styles.add_style('tableSmallCharacter', WD_STYLE_TYPE.CHARACTER)
    tableCharacter.font.name = 'Arial'
    tableCharacter.font.size = Pt(9)
    document.styles['tableSmallCharacter']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')

    # 设置表格字体
    tableCharacter = document.styles.add_style('tableSmallerCharacter', WD_STYLE_TYPE.CHARACTER)
    tableCharacter.font.name = 'Arial'
    tableCharacter.font.size = Pt(8)
    document.styles['tableSmallerCharacter']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')

    # 设置word表格字体
    tableCharacter = document.styles.add_style('tableCharacter', WD_STYLE_TYPE.CHARACTER)
    tableCharacter.font.name = 'Arial'
    tableCharacter.font.size = Pt(10.5)
    document.styles['tableCharacter']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')

    # 正文: 宋体
    style_zero = document.styles.add_style('zero', WD_STYLE_TYPE.CHARACTER)  # 设置一级字样式
    style_zero.font.name = '宋体'
    style_zero.font.size = Pt(12)
    style_zero.font.color.rgb = RGBColor(0x0, 0x0, 0x0)
    document.styles['zero']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')  # 将段落中的所有字体

    # 正文: 宋体
    style_zero = document.styles.add_style('small', WD_STYLE_TYPE.CHARACTER)  # 设置一级字样式
    style_zero.font.name = '宋体'
    style_zero.font.size = Pt(8)
    style_zero.font.color.rgb = RGBColor(0x0, 0x0, 0x0)
    document.styles['small']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')  # 将段落中的所有字体

    # 定义段落样式:首行缩进2个字符、1.5倍行间距
    paragraph_style = document.styles.add_style('paragraphAfterSpace', WD_STYLE_TYPE.PARAGRAPH)
    paragraph_format = paragraph_style.paragraph_format
    paragraph_format.first_line_indent = Pt(24)
    paragraph_format.line_spacing = 1.5
    paragraph_format.space_after = Pt(2)
    paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY


    # 定义段落样式:首行缩进2个字符、1.5倍行间距
    paragraph_style = document.styles.add_style('paragraph', WD_STYLE_TYPE.PARAGRAPH)
    paragraph_format = paragraph_style.paragraph_format
    paragraph_format.first_line_indent = Pt(24)
    paragraph_format.line_spacing = 1.5
    paragraph_format.space_after = Pt(0)
    paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

    # 定义段落样式:无缩进、1.5倍行间距
    paragraph_style = document.styles.add_style('paragraphNoIndent', WD_STYLE_TYPE.PARAGRAPH)
    paragraph_format = paragraph_style.paragraph_format
    paragraph_format.line_spacing = 1.5
    paragraph_format.space_after = Pt(0)
    paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

    # 定义段落样式:靠右对齐
    paragraph_style = document.styles.add_style('paragraphRight', WD_STYLE_TYPE.PARAGRAPH)
    paragraph_format = paragraph_style.paragraph_format
    paragraph_format.line_spacing = 1.5
    paragraph_format.space_after = Pt(0)
    paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
