import pymysql
import traceback

file_path= './data/totalExposureLog/totalExposureLog_1.out'
f = open(file_path)
list = []
for line in f:
    if line in list:
        print(line)
    else:
        list.append(line)
db = pymysql.connect(host="localhost", user="root", password="123456", db="tencent_data", port=3306)
cur = db.cursor()
insert_num = 0
try:
    for line in f:
        line = line.replace("\n", "")
        line = line.replace("\t", ",")
        sql = """INSERT INTO totalexposurelog VALUES (%s);
                """ % line
        cur.execute(sql)
        insert_num += 1
        if insert_num % 100000 == 0:
            print("已导入%d条数据" % insert_num)
    db.commit()
    print("插入成功,已导入" + str(insert_num) + "条数据")
except Exception as e:
    print("出现错误，回滚")
    db.rollback()
    traceback.print_exc()
finally:
    cur.close()
    db.close()
