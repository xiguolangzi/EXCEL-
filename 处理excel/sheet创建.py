
'''
创建sheet方法：
    1.sht = wb.sheets.add()
    2.sht = wb.sheets.add("test",after="sheet2")
查找sheet方法：
    1.sht = wb.sheets("sheet1") #指定名称获取表
    2.sht = wb.sheets(1)    #根据序号获取
    3.sht = wb.sheets.active    #获取当前活动的工作表

'''
import xlwings as xw

# 1.add_book不创建工作部
app = xw.App(visible=True, add_book=False)


# 2.display_alerts参数：打开屏幕中心
app.display_alerts = False
# 3.screen_updating参数：excel提示信息窗口
app.screen_updating = False

filepath = "工作簿3.xlsx"
wb = app.books.open(filepath)

# 4.创建sheet
sht1 = wb.sheets.add()
sht2 = wb.sheets.add()
sht3 = wb.sheets.add("Sheet4",after="Sheet1")
sht4 = wb.sheets.add("Sheet5",before="Sheet2")

print(wb.sheets.count)  #查看表的数量

# 5.打开sheet
sht5 = wb.sheets("Sheet4")
sht5.range("A2").value = "云郎小朋友2"
sht5 = wb.sheets(2)
sht5.range("A3").value = "云郎小朋友3"

wb.save("test1.xlsx")
wb.close()

app.quit()
