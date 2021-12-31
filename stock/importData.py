import pymysql
import traceback
import pandas as pd


# 从数据库中查找数据
def find_data():
    db = pymysql.connect(host="localhost",user="root",password="123456",db="stock",port=3306)
    cur = db.cursor()
    try:
        sql = """select `stock_code`, `stock_name`, `date`, `opening_price`, `top_price`, `floor_price`, `closing_price`, 
        `change_range`, `turnover`, `average_price`, `turnover_rate`, `amplitude`, case `is_harden` when 'Y' then 1 else 0 end , 
        case `is_always_harden` when 'Y' then 1 else 0 end , 
        `morrow_opening_price`,`morrow_top_price`, `morrow_floor_price`, `morrow_closing_price`, `morrow_change_range`, `morrow_turnover`,
         `morrow_average_price`,`morrow_turnover_rate`,`morrow_amplitude`,case `morrow_is_harden` when 'Y' then 1 else 0 end ,
          case `morrow_is_always_harden` when 'Y' then 1 else 0 end ,
          `morrow_open_price_change` from finally_data WHERE (is_always_harden != 'Y' OR morrow_is_always_harden != 'Y')
          AND `morrow_opening_price` is not null"""
        db.commit()
        cur.execute(sql)
        data_tuple = cur.fetchall()
        print("从数据中成功查询出数据")
        return data_tuple
    except Exception as e:
        print("查询数据库，出现错误")
        traceback.print_exc()
    finally:
        cur.close()
        db.close()


def data_to_csv(file, data, columns):
    """
    将数据导出为csv格式
    :param file: 导出的文件地址
    :param data: 文件内容
    :param columns: 文件标题
    :return:
    """
    data = list(data)
    columns = list(columns)
    file_data = pd.DataFrame(data, index=range(len(data)), columns=columns)
    file_data.to_csv(file, index=False, encoding="utf-8", sep="\t")
    print("导出成功")
