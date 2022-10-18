# 1.导入xlwings
import xlwings as xw
'''
# 2.创建app应用对象
app1 = xw.App()
pid = app1.pid
print(pid)

app2 = xw.App()
pid = app2.pid
print(pid)

'''

# 3.查看创建app的个数，即打开excel应用的个数
apps = xw.apps
print(apps.count)

# 4.激活应用
app3 = xw.apps[8444]
app3.activate()

# 5.添加工作簿
wb = app3.books.add()

# 6.设置表页
sht = wb.sheets["sheet1"]

# 7.操作表对应的单元格
sht.range("A1").value = "hahaha"

# 8.保存关闭 工作簿
wb.save()
wb.close()

# 9.关闭应用 kill 和 quit 都可以
app3.kill()
app3.quit()
