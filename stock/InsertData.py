import pymysql
import pandas as pd
import traceback


# 读取CSV文件，返回list
def read_excel_by_pandas(url):
    stock_data = []
    each_stock_data = []
    csv_file = pd.read_csv(url,encoding='utf-8')
    col = len(csv_file.columns)-1
    row = csv_file.shape[0]-2
    for col_index in range(col):
        if col_index < 3:
            continue
        for row_index in range(row):
            if row_index%10 == 0:
                each_stock_data.append(csv_file.iat[row_index,0])  # 股票代码
                each_stock_data.append(csv_file.iat[row_index,1])  # 股票名称
                each_stock_data.append(csv_file.columns[col_index])  # 时间
            #填充其他信息
            value = csv_file.iat[row_index, col_index]
            if (row_index+1)%10 == 0:
                if value =="是":
                    value = "Y"
                elif value == "否":
                    value = "N"
                each_stock_data.append(str(value))
                stock_data.append(each_stock_data)
                each_stock_data = []
            else:
                each_stock_data.append("%.8f" % float(value))
    return stock_data


# 将股票数据插入MySQL数据库
def insert_to_mysql(lt):
    db = pymysql.connect(host="localhost", user="root", password="123456", db="stock", port=3306)
    cur = db.cursor()
    insert_num=0
    try:
        for data in lt:
            param = ""
            for each_data in data:
                if each_data == "nan":param += "null,"
                else:param+="'"+each_data+"',"
            sql = """insert into history (`stock_code`, `stock_name`, `date`,`opening_price`, 
            `top_price`, `floor_price`, `closing_price`,`change_range`, `turnover`, `average_price`, 
            `turnover_rate`, `amplitude`, `is_harden`) values (%s)""" % (param[:-1])
            cur.execute(sql)
            insert_num+=1
        db.commit()
        print("插入成功,已导入" + str(insert_num) + "条数据")
        return insert_num
    except Exception as e:
        print("出现错误，已回滚")
        db.rollback()
        traceback.print_exc()
    finally:
        cur.close()
        db.close()
