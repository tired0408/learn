from stock.InsertData import *

total_num = 0
for i in range(1,10):
    total_num+=insert_to_mysql(read_excel_by_pandas(url = "D:\\胡\\桌面文件\\股票数据\\a"+str(i)+".csv"))
    print("a"+str(i)+".csv文件已导入数据库")
print("总共插入了"+str(total_num)+"条数据")