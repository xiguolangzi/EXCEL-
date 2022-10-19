'''
介绍单元格的引用区域操作：
    1.多个单元格的选择
    2.单元格的便宜与扩展
    3.区域单元格进行数据存储

'''

import xlwings as xw

app = xw.App(visible=True, add_book=False)
app.display_alerts = False
app.screen_updating = False

filepath = "test1.xlsx"
wb = app.books.open(filepath)

sht1 = wb.sheets["Sheet1"]

# 1.单个单元格，单个值
sht1.range("A8").value = "|有准备才有机会"
# 2.单个单元格，A9右偏移2个值进行填充
sht1.range("A9").value = ["你好","我是","小可爱"]

# 3.扩展区域赋值，C2向下继续填充单元格
sht1.range("C2").options(transpose=True).value = ["AAA","BBB","CCC","DDD"]

# 4.down 下方选中\right 右方选中\table 又下方选中(1,2)为单位的行，向下选中填充
sht1.range("D8").options(expand="table").value = [(1,2),(3,4)]

wb.save()
wb.close()
app.quit()


