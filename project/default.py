from project.utils import set_cell_border


# 设置表格样式
def set_table_border(table):
    for cell in table.rows[0].cells:
        set_cell_border(cell, top={"sz": 12, "val": "single", "space": "0"})
    for cell in table.rows[-1].cells:
        set_cell_border(cell, bottom={"sz": 12, "val": "single", "space": "0"})
    for row in table.rows[0:len(table.rows) - 1]:
        for cell in row.cells:
            set_cell_border(cell, bottom={"sz": 6, "val": "dotted", "space": "0"})
    for key, column in enumerate(table.columns):
        if key == len(table.columns) - 1:
            continue
        for cell in column.cells:
            set_cell_border(cell, right={"sz": 6, "val": "dotted", "space": "0"})

# 设置段落样式
