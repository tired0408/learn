import time
import collections
import numpy as np
import pandas as pd
import xlwings as xw
from itertools import combinations
from typing import List, Dict
from difflib import SequenceMatcher


class Unit:
    """每个单元类"""

    def __init__(self) -> None:
        self.index: List[int] = []
        self.amount: List[int] = []


def similarity_match(a, b):
    """定义相似度匹配函数"""
    similarity = SequenceMatcher(None, a, b).ratio()
    return similarity >= 0.8


def get_database(path):
    """"获取各个终端客户的产品名称转化为标准名称的数据库"""
    rd = {}
    df = pd.read_excel(path, header=1)
    for row in df.iterrows():
        rd[row["合并项"]] = row["品规(清洗后)"]
    return rd


def get_zgzy_data(path):
    """获取中国中药的数据及客户名称转化为标注名称的数据库"""
    rd: Dict[str, Dict[str, Unit]] = {}
    name2standard: Dict[str, Dict[str, str]] = {}
    df = pd.read_excel(path, header=0)
    df = df[["销售日期", "购入客户名称(原始)", "购入客户名称(清洗后)", "品规(清洗后)", "批次", "数量"]]
    for row in df.iterrows():
        name = row["购入客户名称(原始)"]
        name_standard = row["购入客户名称(清洗后)"]
        key = "@@".join([row["销售日期", "品规(清洗后)", "批次"]])

        name2standard[key][name] = name_standard

        if key not in rd:
            rd[key] = {name_standard: Unit()}
        if key not in rd[key]:
            rd[key][name_standard] = Unit()
        rd[key][name_standard].amount.append(row["数量"])
        rd[key][name_standard].index.append(row.index)

    return rd, name2standard


def get_client_data(path, product_database, client_database: Dict[str, Dict[str, str]]) -> Dict[str, Dict[str, Unit]]:
    """整理终端客户的数据"""
    rd: Dict[str, Dict[str, Unit]] = {}
    df = pd.read_excel(path, header=0)
    df = df[["销售日期", "客户名称", "商品名称", "规格", "批号", "销售数量"]]
    for row in df.iterrows():
        quality_regulation = row["商品名称"] + row["规格"]
        if quality_regulation not in product_database:
            print(f"[警告]{quality_regulation}并未在产品库中，请核实")
            continue
        quality_regulation = product_database[quality_regulation]

        key = "@@".join([row["销售日期", quality_regulation, "批号"]])

        name_standard = row["客户名称"]
        if name_standard in client_database[key]:
            name_standard = client_database[key][name_standard]
        else:
            similarity_list = []
            for name in client_database[key].values():
                if not similarity_match(name, name_standard):
                    continue
                similarity_list.append(name)
            if len(similarity_list) != 1:
                print(f"[警告]该客户名称未找到匹配项，请核实:{key}, {name_standard}, {similarity_list}")
            else:
                name_standard = similarity_list[0]

        if key not in rd:
            rd[key] = {name_standard: Unit()}
        if key not in rd[key]:
            rd[key][name_standard] = Unit()
        rd[key][name_standard].amount.append(row["数量"])
        rd[key][name_standard].index.append(row.index)
    return rd


def find_combination(nums, target):
    """找到nums中所有可能的组合,其和等于target"""
    for r in range(1, len(nums) + 1):
        for combo in combinations(nums, r):
            if sum(combo) == target:
                return combo
    return None


def match_and_diff(list1: List, list2: List):
    """匹配不同的值"""
    list1 = list1.copy()
    list2 = list2.copy()
    unmatched_list1 = []
    unmatched_list2 = []

    while len(list1) != 0 and len(list2) != 0:
        max1 = max(list1)
        max2 = max(list2)
        if max1 == max2:
            list1.remove(max1)
            list2.remove(max2)
            continue
        if max1 > max2:
            list1.remove(max1)
            combo_for_max1 = find_combination(list2, max1)
            if combo_for_max1:
                for num in combo_for_max1:
                    list2.remove(num)
            else:
                unmatched_list1.append(max1)
        else:
            list2.remove(max2)
            combo_for_max2 = find_combination(list1, max2)
            if combo_for_max2:
                for num in combo_for_max2:
                    list1.remove(num)
            else:
                unmatched_list2.append(max2)
    unmatched_list1.extend(list1)
    unmatched_list2.extend(list2)
    return unmatched_list1, unmatched_list2


def main():
    zgzy_path = r"E:\NewFolder\liuxiang\商业流向明细报表 湖南 22年7月-24年5月-国控岳阳.xlsx"
    client_path = r"E:\NewFolder\liuxiang\国药控股岳阳有限公司.xls"
    database_path = r"E:\NewFolder\liuxiang\湖南 商业品规清洗桥梁表（使用）.xlsx"
    product_database = get_database(database_path)
    zgzy_dict, client_database = get_zgzy_data(zgzy_path)
    client_dict = get_client_data(client_path, product_database, client_database)
    # 比对中国中药及终端客户的数据差异
    zgzy_different = []
    client_different = []
    for client_k1, client_v1 in client_dict.items():
        if client_k1 not in zgzy_dict:
            for _, client_v2 in client_v1.items():
                client_different.extend(client_v2.index)
            continue
        zgzy_v1 = zgzy_dict[client_k1]
        for client_k2, client_v2 in client_v1.items():
            if client_k2 not in zgzy_v1:
                client_different.extend(client_v2.index)
                continue
            zgzy_v2 = zgzy_v1.pop(client_k2)
            unmatched_zgzy, unmatched_client = match_and_diff(zgzy_v2.amount, client_v2.amount)
            for amount in unmatched_zgzy:
                zgzy_different.append(zgzy_v2.index[zgzy_v2.amount.index(amount)])
            for amount in unmatched_client:
                client_different.append(client_v2.index[client_v2.amount.index(amount)])
    # 补充中国中药里面的差异数据
    for k1, v1 in zgzy_dict.items():
        for k2, v2 in v1.items():
            zgzy_different.extend(v2.index)
    print("程序已运行完毕")
    print(zgzy_different)
    print(client_different)


if __name__ == '__main__':
    main()
