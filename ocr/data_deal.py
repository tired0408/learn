from glob import glob
import os
import numpy as np
import shutil
import pandas as pd
import re
import random
import math
# *******************************
# ********** 标注之前需要处理的一些方法
# *******************************
from glob import glob
import os
import shutil
import pandas as pd
import re
import random
import math

def generate_label(path):
    """
    根据文件名生成标签
    :return:
    """

    path_list = glob(os.path.join(path, "*.jpg"))
    data_dict = {"name": path_list, "label":[""]*len(path_list)}
    data = pd.DataFrame(data_dict)
    data = data.astype("str")
    data["name"] = data["name"].apply(lambda x:os.path.basename(x))
    data.to_csv(os.path.join(path, "label.txt"), index=False, header=False, encoding="gbk", sep=" ")
def extract_by_file(path):
    """
    根据文件提取标签
    :return:
    """
    labels = pd.read_csv(os.path.join(path, "label.txt"), sep=" ", encoding="gbk", header=None)
    jpg_path_list = pd.Series(os.listdir(path))
    jpg_path_list = jpg_path_list[jpg_path_list.str.contains(".jpg")].tolist()
    labels = labels[labels[0].isin(jpg_path_list)]
    labels.to_csv(os.path.join(path, "label.txt"), sep=" ", header=None, index=False, encoding="gbk")
def label_to_filename(path):
    """
    将文件名修改为标签，如果为空，命名为NONE，并修改label.txt文件
    :return:
    """
    labels = pd.read_csv(os.path.join(path, "label.txt"), sep=" ", encoding="gbk", header=None)
    for index, row in labels.iterrows():
        if row.isna()[1]:
            new_name = "NONE_%d.jpg" % index
        else:
            new_name = "%s_%d.jpg" % (self.address_list(r"[.|]*", "", row[1]), index)
        os.rename(os.path.join(path, row[0]), os.path.join(path, new_name))
        labels.loc[index, 0] = new_name
    labels.to_csv(os.path.join(path, "label.txt"), sep=" ", header=None, index=False, encoding="gbk")
    print("已将文件名全部修改为标签名")
def merge_mul_file(path, save_path, has_history=False):
    """
    将多个数据文件合并到一起，并将文件名改为标签名，空标签命名为None
    :param has_history: 是否记录变化的历史情况，你们不用管
    :param path:原始地址
    :param save_path: 保存地址
    :return:
    """
    file_name = os.listdir(path)
    if not os.path.exists(save_path):
        os.mkdir(save_path)
    save_labels = pd.DataFrame()
    history_df = pd.DataFrame()
    i = 0
    for fname in file_name:
        fpath = os.path.join(path, fname)
        jpg_path = glob(os.path.join(fpath, "*.jpg"))
        if len(jpg_path)== 0:
            continue
        flabels_path = os.path.join(fpath, "label.txt")
        if os.path.exists(flabels_path):
            flabels = pd.read_csv(flabels_path, sep=" ", encoding="gbk", header=None)
        else:
            data_dict = {0: jpg_path, 1: [""] * len(jpg_path)}
            flabels = pd.DataFrame(data_dict)
            flabels = flabels.astype("str")
            flabels[0] = flabels[0].apply(lambda x: os.path.basename(x))
        for index,row in flabels.iterrows():
            if not row.isna()[1] and re.search('[^a-zA-Z0-9|]', row[1]) is not None:
                print("出现错误标注，含有其他字符，文件夹：%s，文件名：%s，标签名：%s" % (fname, row[0], row[1]))
                continue
            i += 1
            if row.isna()[1]:
                new_name = "NONE_%d.jpg" % i
            else:
                new_name = "%s_%d.jpg" % (re.sub(r"[.|]*", "", row[1]), i)
            shutil.copy(os.path.join(fpath, row[0]), os.path.join(save_path, new_name))
            if has_history:
                his_row = row.append(pd.Series([fname, new_name]), ignore_index=True)
            row[0] = new_name
            save_labels = save_labels.append(row)
            if has_history:
                his_row.name = row.name
                history_df = history_df.append(his_row)
    save_labels.to_csv(os.path.join(save_path, "label.txt"), sep=" ", encoding="gbk", header=False, index=False)
    if has_history:
        history_df.to_csv(os.path.join(save_path, "history.txt"), sep=" ", encoding="gbk", header=False, index=False)
    print("样本总数：", save_labels.shape[0])
def view_the_most_label(max_num=50):
    """
    查看当前已有许多样本的标签
    :return:
    """
    path = r"E:\work\09data\sample"
    file_name = os.listdir(path)
    labels = pd.DataFrame()
    for fname in file_name:
        fpath = os.path.join(path, fname)
        flabels_path = os.path.join(fpath, "label.txt")
        flabels = pd.read_csv(flabels_path, sep=" ", encoding="gbk", header=None)
        labels = labels.append(flabels)
    labels[1] = labels[1].apply(lambda x: re.sub(r'[.|]*', "", x))
    label_count = labels[1].value_counts()
    print(label_count[label_count > max_num])
    ignore_label = label_count[label_count > max_num].index.tolist()
    print("忽视的label名单：", ignore_label)
    return ignore_label
def view_data_label_situation(path, max_num=10):
    """
    查看当前数据集的百度API识别情况
    :return:
    """
    labels = pd.read_csv(os.path.join(path, "label.txt"), sep=" ", encoding="gbk", header=None)
    label_count = labels[1].value_counts()
    print(label_count[label_count > max_num])
    ignore_label = label_count[label_count > max_num].index.tolist()
    return ignore_label
def del_file_by_label(path, ignore_label, ignore_contains_label, regular_list, is_save_del=True):
    """
    根据忽视标签列表删除文件
    :param is_save_del:
    :param regular_list:
    :param ignore_contains_label:
    :param path:
    :param ignore_label:
    :return:
    """
    labels = pd.read_csv(os.path.join(path, "label.txt"), sep=" ", header=None, encoding="gbk")
    na_list = labels[labels[1].isna()]
    labels = labels.dropna(axis=0, how="any")  # 去除有缺失值的行
    row_num = labels.shape[0]
    # 如果要保存删减图片，则定义保存地址
    if is_save_del:
        save_path = path+"_deldata"
        if os.path.exists(save_path):
            shutil.rmtree(save_path)
        os.mkdir(save_path)
    # 删除忽视标签里面的文件
    res = labels[1].isin(ignore_label)
    for dfile in labels[res][0].tolist():
        if is_save_del:
            shutil.move(os.path.join(path, dfile), os.path.join(save_path, dfile))
        else:
            os.remove(os.path.join(path, dfile))
    labels = labels[~res]
    # 删除标签含有列表内字符的文件
    for icl in ignore_contains_label:
        res = labels[1].str.contains(icl)
        for dfile in labels[res][0].tolist():
            if is_save_del:
                shutil.move(os.path.join(path, dfile), os.path.join(save_path, dfile))
            else:
                os.remove(os.path.join(path, dfile))
        labels = labels[~res]
    # 根据正则表达式删除文件
    for rl in regular_list:
        res = labels[1].str.contains(rl, regex=True)
        for dfile in labels[res][0].tolist():
            if is_save_del:
                shutil.move(os.path.join(path, dfile), os.path.join(save_path, dfile))
            else:
                os.remove(os.path.join(path, dfile))
        labels = labels[~res]
    labels = labels.append(na_list)
    final_num = labels.shape[0]
    print("原始数量：%d，删减后数量：%d" % (row_num, final_num))
    labels.to_csv(os.path.join(path, "label.txt"), sep=" ", header=None, index=False, encoding="gbk")
def split_data(path):
    """
    将需要标注的数据分割成2000个一份
    :return:
    """
    save_path = path + "_deal"
    if not os.path.exists(save_path):
        os.mkdir(save_path)
    labels = pd.read_csv(os.path.join(path, "label.txt"), sep=" ", encoding="gbk", header=None)
    each_file_num = 2000
    file_num = math.ceil(labels.shape[0]/each_file_num)
    for i in range(file_num):
        save_labels = labels.iloc[i*each_file_num:(i+1)*each_file_num]
        each_save_path = os.path.join(save_path, "%02d" % i)
        os.mkdir(each_save_path)
        for index, row in save_labels.iterrows():
            shutil.copy(os.path.join(path, row[0]), os.path.join(each_save_path, row[0]))
        save_labels.to_csv(os.path.join(each_save_path, "label.txt"), index=False, header=False, encoding="gbk", sep=" ")


def view_image_by_label(path, view_label = {}):
    """
    根据列表拷贝出图片，查看图片情况
    :param view_label:
    :param path:
    :return:
    """
    labels = pd.read_csv(os.path.join(path, "label.txt"), sep=" ", header=None, encoding="gbk")
    save_path = path + "_deal"
    if os.path.exists(save_path):
        shutil.rmtree(save_path)
    os.mkdir(save_path)
    labels = labels[labels[1].isin(view_label)]
    # labels = labels[labels[1].str.contains(r'KG',regex=False)]
    for index, row in labels.iterrows():
        shutil.copy(os.path.join(path, row[0]), os.path.join(save_path, row[0]))

def check_data(path):
    """
    校验label.txt是否与文件一一对应
    :param path:
    :return:
    """
    labels = pd.read_csv(os.path.join(path, "label.txt"), sep=" ", header=None, encoding="gbk")
    label_file_name = labels[0].tolist()
    jpg_path_list = pd.Series(os.listdir(path))
    jpg_path_list = jpg_path_list[jpg_path_list.str.contains(".jpg")].tolist()
    print("label.txt中有该文件，但文件夹中没有：")
    print(set(label_file_name).difference(set(jpg_path_list)))
    print("文件夹中有该文件，但label.txt中没有：")
    print(set(jpg_path_list).difference(set(label_file_name)))

def del_empty_label(path):
    """
    删除空标签的文件
    :param path:
    :return:
    """
    labels = pd.read_csv(os.path.join(path, "label.txt"), sep=" ", header=None, encoding="gbk")
    to_del_df = labels[labels[1].isna()]
    to_save_df = labels[~labels[1].isna()]
    to_save_df.to_csv(os.path.join(path, "label.txt"), index=False, header=False, encoding="gbk", sep=" ")
    print("the del of empty label has been finnished.")

def post_process(path):
    """
    数据标注的后处理工作，1、删除标签为空的文件。2、合并多个文件夹、并将文件名改为标签
    :return:
    """
    file_list = os.listdir(path)
    for f_name in file_list:
        del_empty_label(os.path.join(path, f_name))
    merge_mul_file(path, path, has_history=True)
    for f_name in file_list:
        shutil.rmtree(os.path.join(path, f_name))
    print("The data which has been labeled are sorted out.")
# 预处理：
# 使用百度先预先跑一遍数据
# 大致看下各个文件夹的情况，如果选择需要的留下，并进行合并
# 将文件名命名为标签
# 通过标签过滤一些已有足够样本的，和一些不需要的样本
# 人为通过大小排序筛选掉一些数据，根据文件名提取标签
# 人为通过名称排序筛选掉一些数据，并记录没用数据可能会识别成的标签名，根据文件名提取标签
# 将数据分成2000一份
# 后处理：
# 将文件夹进行合并
# 后处理：
# 删除标签为空的文件
# 合并多个文件夹
# 将文件名改为标签
can_execute = True
if can_execute:
    f_path = r"E:\work\09data\to_label\20191206\all"
    # merge_mul_file(r"E:\work\09data\sample", r"E:\work\09data\sample\all")
    # label_to_filename(f_path)
    ignore_contains_list = ["LBS", "KG", "CUM", "CAP", "AUTION", "MAX", "CUFT", "AERSK", "CARGO"]
    ignore_list = ['45G1', 'TARE', 'NET', 'CAUTION', 'PAYLOAD', 'HIGH', 'CUBE', 'MAERSK', '22G1', 'TRITON']
    ignore_list.extend(["45G", "220", "ANIONCDGISIICS", "ANIONGDCISIICS", "ANIONCDCISIICS", "22GI", "450", "456",
                   "BEACON", "FLOREN", "LOREN", "MAXCARGO"])
    ignore_list.extend(['LORENS', 'TAPE', 'LOREN', 'BEACON', 'ARE', 'FLOREN', 'IARE', 'ANIONCDCISIICS', '2201',
                        '226', 'MAXWI', 'FLORENS', 'TAR', 'UCAP', '456', '22GI', '450', 'MAXCARGO', '2CUM', 'NEI', '764CUM'])
    ignore_list.extend(["2261", "4561", "TAREWT", "BROMMA", "CONTAINER", "4567", "MAXWT", "45G7", "45GI", "45GT", "45G", "A5G","A5G|",
                        "ASG", "ASG|","45G|", "456|", "456|", "5G1"
                        "22G7", "RITON", "GROSS", "MGW","CMACGM", "CATRAON", "ANIONCIDGISIICS"])
    regu_list = [r'[a-z]{2,}', r'^[a-zA-Z0-9|]{1,2}$',r"^BRO", r"^[0-9]+[a-zA-Z]+", r"^[0-9]{3}$"]
    # ignore_list = view_the_most_label()
    # del_file_by_label(f_path, set(ignore_list), ignore_contains_list, regu_list, is_save_del=True)
    # view_list = view_data_label_situation(f_path, max_num=10)
    # view_list = set(view_list).difference({"ZIM", "42G1", "4261", "TEMU", "TCNU", "TCLU", "RSK", "MSKU","MRKU", "L5G1"
    #                                        "MAEU", "HMMU", "HMM", "MAEU", "L5G1", "BMOU"})
    # view_image_by_label(f_path, view_list)
    # generate_label(f_path)

    # split_data(r"E:\work\09data\to_label\20191206\all")
    # generate_label(r"E:\work\09data\to_label\20191206\test")
    # post_process(r"E:\work\09data\sample\20191207")
    # label_to_filename(r"E:\work\09data\sample\20191207")
    check_data(r"E:\work\09data\sample\20191207")
    extract_by_file(r"E:\work\09data\sample\20191207")
else:
    del_file_by_label(r"E:\work\09data\to_label\20191206\02", [], [])
    generate_label(r"E:\work\09data\to_label\20191206\test")
    label_to_filename(r"E:\work\09data\to_label\20191206\01")
    merge_mul_file(r"E:\work\09data\to_label\20191206\01")
    extract_by_file()
    check_data()
    split_data()
    del_empty_label()
