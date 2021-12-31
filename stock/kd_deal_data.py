import pandas as pd
import time
import csv
import os


def extracted_features(date_frame, deal_titles):
    """
    处理特征
    :param date_frame:
    :param deal_titles:
    :return:
    """
    all_features = [{} for i in deal_titles]
    data_max, data_min, data_sum, data_num = 0, 99, 0, 0
    for row_tuple in date_frame.iterrows():
        # 处理时间特征
        time_hour = time.localtime(row_tuple[1]["time"]).tm_hour+1  # 通过时间戳，获取时间小时数
        data_num += 1
        data_sum += time_hour
        if data_max < time_hour:
            data_max = time_hour
        if data_min > time_hour:
            data_min = time_hour
        # 处理其他特征
        for i in range(len(deal_titles)):
            title = deal_titles[i]
            val = row_tuple[1][title]
            if val is None:
                val = title+"_break"
            if val in all_features[i]:
                all_features[i][val] += 1
            else:
                all_features[i][val] = 1
    print("特征处理完毕，获得特征及时间的相关信息。")
    return [{"max_minus_min": data_max-data_min, "min": data_min, "average": data_sum / data_num}, all_features]


# 根据细分特征，构建标题
def build_title(input_data):
    finally_title = ["time"]
    for i in range(len(input_data)):
        features_content = input_data[i]
        for key, value in features_content.items():
            if value > 10:
                finally_title.append(key)
        finally_title.append(features_list[i]+"_other")
    print("构建标题成功，标题长为：", len(finally_title))
    return finally_title


def deal_data(date_frame, titles, time_info, features):
    """
    根据构建好的标题，处理数据，导出CSV文件
    :param date_frame: 需要处理的数据（dateframe格式）
    :param titles: 处理好的特征标题
    :param time_info: 时间特征的信息，包含最大值减最小值，平均值
    :param features: 筛选出来的特征
    :return:
    """
    save_file_path = "E:\\桌面文件\\算法比赛题\\finally_result"
    # 清空该文件夹下的问题
    for i in os.listdir(save_file_path):
        path_file = os.path.join(save_file_path, i)
        if os.path.isfile(path_file):
            os.remove(path_file)
    # 创建文件并写入标题
    csvfile = open(save_file_path+"\\finally_test_data%d.txt" % (round(time.time())), "w", newline="",
                   encoding="utf-8")
    csvfile2 = open(save_file_path+"\\finally_test_label%d.txt" % (round(time.time())), "w", newline="",
                    encoding="utf-8")
    writer = csv.writer(csvfile, delimiter="\t")
    writer.writerow(titles)  # 写入标题
    writer2 = csv.writer(csvfile2, delimiter="\t")
    writer2.writerow(["click"])
    n = 0
    file_num = 1
    # 写入CSV文件内容
    for row_tuple in date_frame.iterrows():  # 遍历数据
        time_hour = time.localtime(row_tuple[1]["time"]).tm_hour + 1  # 时间转化为小时
        each_data = [0 for i in range(len(titles))]  # 存储处理后的每行数据
        # 归一化时间数据
        if time_info["max_minus_min"] != 0:
            each_data[0] = (time_hour-time_info["min"])/time_info["max_minus_min"]
        # 处理其他特征
        for feature in features:
            value = row_tuple[1][feature]  # 获取大特征中的值
            if value is None:
                value = feature+"_break"
            if value in titles:
                each_data[titles.index(value)] = 1
            else:
                each_data[titles.index(feature+"_other")] = 1
        writer.writerow(each_data)  # 一行一行写入数据
        writer2.writerow([row_tuple[1]["click"]])
        n += 1
        if n % 10000 == 0:
            print("已写入", n, "条数据")
        if n % 100000 == 0:
            file_num += 1
            csvfile.close()
            csvfile2.close()
            csvfile = open(save_file_path + "\\finally_test_data%d.txt" % (round(time.time())), "w", newline="",
                           encoding="utf-8")
            csvfile2 = open(save_file_path + "\\finally_test_label%d.txt" % (round(time.time())), "w", newline="",
                            encoding="utf-8")
            writer = csv.writer(csvfile, delimiter="\t")
            writer.writerow(titles)  # 写入标题
            writer2 = csv.writer(csvfile2, delimiter="\t")
            writer2.writerow(["click"])
    print("已导出%d个CSV文件，写入%d行数据" % (2*file_num, n))
    csvfile.close()
    csvfile2.close()


path = "E:\\桌面文件\\算法比赛题\\test.txt"
path = "E:\\桌面文件\\算法比赛题\\round1_iflyad_train.txt"
data = pd.read_csv(path, delimiter="\t")
print("原数据有:", data.shape[0], "条")
# print(data.columns)  # 打印标题
features_list = ["city", "osv", "make", "inner_slot_id", "app_id", "adid", "f_channel", "app_cate_id",
                   "province", "advert_id", "advert_industry_inner"]
features_info = extracted_features(data, features_list)
deal_data(data, build_title(features_info[1]), features_info[0], features_list)
