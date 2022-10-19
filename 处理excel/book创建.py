import xlwings as xw

'''
创建book的方法：
    1.创建：xw.Book()  xw.books.add()
    2.指定名称：xw.Book("Book01")    xw.books["Book01","Book02",...]
    3.执行路径文件：xw.Book(r"C:/path/to/file.xlsx")   xw.books.open(r"C:/path/to/file.xlsx")
    

'''
# 1.打开 工作簿，要使用全名，否则会报fullname
book = xw.Book()
# 2.打开表页
sht = book.sheets["Sheet1"]
# 3.输入单元格内容
sht.range("A3").value = "hello office"
sht.range("B3").value = "hello office2"

# 4.保存 book.save(） , 创建并保存 book.save("工作簿1.xlsx")
book.save("工作簿1.xlsx")
# 5.关闭
book.close()