import os
import xlwt
import datetime
import collections
import pandas as pd
from xlwt.Worksheet import Worksheet, Style
from itertools import combinations
from typing import List, Dict, Tuple
from difflib import SequenceMatcher


class Unit:
    """每个单元类"""

    def __init__(self) -> None:
        self.index: List[int] = []
        self.amount: List[int] = []


def similarity_match(a, b):
    """定义相似度匹配函数"""
    similarity = SequenceMatcher(None, a, b).ratio()
    return similarity


def get_database(path):
    """"获取各个终端客户的产品名称转化为标准名称的数据库"""
    rd = {}
    df = pd.read_excel(path, header=1)
    for _, row in df.iterrows():
        rd[row["合并项"]] = row["品规(清洗后)"]
    return rd


def get_zgzy_data(path) -> Tuple[pd.DataFrame, pd.DataFrame, Dict[str, Dict[str, Unit]], Dict[str, str]]:
    """获取中国中药的数据及客户名称转化为标注名称的数据库"""
    rd: Dict[str, Dict[str, Unit]] = {}
    name2standard: Dict[str, Dict[str, str]] = {}

    raw_df = pd.read_excel(path, header=0)
    raw_df["序号"] = range(1, len(raw_df) + 1)

    df = raw_df[["销售日期", "购入客户名称(原始)", "购入客户名称(清洗后)", "品规(清洗后)", "标准批号", "数量"]]
    df.loc[:, "标准批号"] = df["标准批号"].astype(str)
    df = df.apply(lambda x: x.str.replace(' ', '') if x.dtype == 'object' else x)

    for index, row in df.iterrows():
        name = row["购入客户名称(原始)"]
        name_standard = row["购入客户名称(清洗后)"]
        amount = row["数量"]

        key: str = "@@".join([row["销售日期"], row["品规(清洗后)"], row["标准批号"]])

        if key not in name2standard:
            name2standard[key] = {}
        name2standard[key][name] = name_standard

        if key not in rd:
            rd[key] = {name_standard: Unit()}
        if name_standard not in rd[key]:
            rd[key][name_standard] = Unit()
        rd[key][name_standard].amount.append(amount)
        rd[key][name_standard].index.append(index)

    return_df = raw_df.reindex(columns=["销售日期", "购入客户名称(原始)", "品规(清洗后)", "标准批号", "数量", "是否已红冲", "序号"])
    return_df = return_df.rename(columns={
        "购入客户名称(原始)": "客户名称",
        "标准批号": "批号",
        "数量": "销售数量"
    })
    return_df["来源"] = ["中国中药表"] * len(return_df)
    return raw_df, return_df, rd, name2standard


def get_client_data(path, client_database: Dict[str, Dict[str, str]]
                    ) -> Tuple[pd.DataFrame, pd.DataFrame, Dict[str, Dict[str, Unit]]]:
    """整理终端客户的数据"""
    rd: Dict[str, Dict[str, Unit]] = {}

    raw_df = pd.read_excel(path, header=0)
    raw_df["销售日期"] = pd.to_datetime(raw_df['销售日期']).dt.strftime('%Y-%m-%d')
    raw_df["序号"] = [f"({i})" for i in range(1, len(raw_df) + 1)]

    df = raw_df[["销售日期", "客户名称", "品规(清洗后)", "批号", "销售数量"]]
    df.loc[:, "批号"] = df["批号"].astype(str)
    df = df.apply(lambda x: x.str.replace(' ', '') if x.dtype == 'object' else x)

    for index, row in df.iterrows():
        quality_regulation = row["品规(清洗后)"]
        date: pd.Timestamp = row["销售日期"]

        key: str = "@@".join([date, quality_regulation, row["批号"]])

        name_standard = row["客户名称"]
        if key not in client_database:
            pass
        elif name_standard in client_database[key]:
            name_standard = client_database[key][name_standard]
        else:
            similarity_score = 0
            similarity_name = None
            name_list = list(client_database[key].values())
            for name in name_list:
                score = similarity_match(name, name_standard)
                if score < 0.8 or score < similarity_score:
                    continue
                similarity_score = score
                similarity_name = name
            if similarity_name is None:
                print(f"[警告]客户名匹配异常, 未找到匹配项，请核实:{key}, {name_standard}, {name_list}")
            else:
                name_standard = similarity_name
        if key not in rd:
            rd[key] = {name_standard: Unit()}
        if name_standard not in rd[key]:
            rd[key][name_standard] = Unit()
        rd[key][name_standard].amount.append(row["销售数量"])
        rd[key][name_standard].index.append(index)

    return_df = raw_df.reindex(columns=["销售日期", "客户名称", "品规(清洗后)", "批号", "销售数量", "是否已红冲", "序号"])
    return_df["来源"] = ["客户表"] * len(return_df)
    return raw_df, return_df, rd


def compare_different(zgzy_dict: Dict[str, Dict[str, Unit]], client_dict: Dict[str, Dict[str, Unit]]):
    """比对不一致的数据"""
    zgzy_different = []
    client_different = []
    for client_k1, client_v1 in client_dict.items():
        if client_k1 not in zgzy_dict.keys():
            for _, client_v2 in client_v1.items():
                client_different.extend(client_v2.index)
            continue
        zgzy_v1 = zgzy_dict[client_k1]
        for client_k2, client_v2 in client_v1.items():
            if client_k2 not in zgzy_v1.keys():
                client_different.extend(client_v2.index)
                continue
            zgzy_v2 = zgzy_v1.pop(client_k2)
            unmatched_zgzy, unmatched_client = match_and_diff(zgzy_v2.amount, client_v2.amount)
            for amount in unmatched_zgzy:
                amount_index = zgzy_v2.amount.index(amount)
                zgzy_v2.amount.pop(amount_index)
                zgzy_different.append(zgzy_v2.index.pop(amount_index))
            for amount in unmatched_client:
                amount_index = client_v2.amount.index(amount)
                client_v2.amount.pop(amount_index)
                client_different.append(client_v2.index.pop(amount_index))
    # 补充中国中药里面的差异数据
    for _, v1 in zgzy_dict.items():
        for _, v2 in v1.items():
            zgzy_different.extend(v2.index)
    return zgzy_different, client_different


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


def get_xlwt_color_style(color):
    """获取颜色样式"""
    style = xlwt.XFStyle()
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = xlwt.Style.colour_map[color]
    style.pattern = pattern
    return style


def deal_excel(data: pd.DataFrame, ws: Worksheet, indexs):
    """根据索引填充颜色, 并新增数据"""
    # 新增数据
    data = data.fillna("")
    data["年月"] = pd.to_datetime(data['销售日期']).dt.strftime('%Y-%m')
    data["年月+品规"] = data["品规(清洗后)"] + data["年月"]
    # 修改类型
    data = data.astype(str)
    # 写入数据
    data["color"] = ""
    data.loc[indexs, "color"] = "red"
    write_data(ws, data)


def statistics_names(zgzy_df_tidy: pd.DataFrame, client_df_tidy: pd.DataFrame):
    """将数据按品规的方式进行统计"""
    quality_dict = collections.defaultdict(lambda: [0, 0, set()])
    month_quality_dict = collections.defaultdict(lambda: [0, 0])
    for i, df in enumerate([zgzy_df_tidy, client_df_tidy]):
        for _, row in df.iterrows():
            name: str = row["品规(清洗后)"]
            name = name.replace(" ", "")
            date = row["销售日期"][:7]
            amount = row["销售数量"]
            quality_dict[name][i] += amount
            quality_dict[name][2].add(date)
            month_quality_dict[f"{name}{date}"][i] += amount

    quality_df = []
    for key, value in quality_dict.items():
        quality_df.append([key, value[0], value[1], list(value[2])])
    quality_df = pd.DataFrame(quality_df, columns=["品规(清洗后)", "系统数量", "商业数量", "年月"])

    month_quality_df = []
    for key, value in month_quality_dict.items():
        month_quality_df.append([key, value[0], value[1]])
    month_quality_df = pd.DataFrame(month_quality_df, columns=["年月+品规", "系统数量", "商业数量"])

    return quality_df, month_quality_df


STR2COLOR = {
    "yellow": get_xlwt_color_style("yellow"),
    "orange": get_xlwt_color_style("orange"),
    "red": get_xlwt_color_style("red"),
    "": Style.default_style
}


def tidy_different_data(zgzy_df_tidy: pd.DataFrame, zgzy_different,
                        client_df_tidy: pd.DataFrame, client_different):
    """整理差异数据表"""
    zgzy_df_different = zgzy_df_tidy.copy()
    zgzy_df_different["color"] = ""
    zgzy_df_different.loc[zgzy_different, "color"] = "yellow"
    client_df_different = client_df_tidy.copy()
    client_df_different["color"] = ""
    client_df_different.loc[client_different, "color"] = "orange"
    df = pd.concat([zgzy_df_different, client_df_different], ignore_index=True)
    df = df.fillna("")
    return df


def tidy_different_detail(quality_df: pd.DataFrame, month_quality_df: pd.DataFrame):
    """整理差异明细"""
    quality_df_len = len(quality_df)
    month_quality_df_len = len(month_quality_df)

    titles = ["品规(清洗后)", "系统数量", "商业数量", "差异", "备注", "", "", "年月+品规", "系统数量", "商业数量", "差异", "备注"]
    data = pd.DataFrame("", index=range(max(quality_df_len, month_quality_df_len)), columns=titles)

    data.iloc[:quality_df_len, 0] = quality_df["品规(清洗后)"]
    data.iloc[:quality_df_len, 1] = quality_df["系统数量"]
    data.iloc[:quality_df_len, 2] = quality_df["商业数量"]
    data.iloc[:quality_df_len, 3] = quality_df["系统数量"] - quality_df["商业数量"]

    data.iloc[:month_quality_df_len, 7] = month_quality_df["年月+品规"]
    data.iloc[:month_quality_df_len, 8] = month_quality_df["系统数量"]
    data.iloc[:month_quality_df_len, 9] = month_quality_df["商业数量"]
    data.iloc[:month_quality_df_len, 10] = month_quality_df["系统数量"] - month_quality_df["商业数量"]
    return data


def tidy_zgzy_perspective(quality_df: pd.DataFrame, month_quality_df: pd.DataFrame):
    """整理透视表数据"""
    # 清理非中国中药的有效数据
    quality_df = quality_df[quality_df["系统数量"] != 0]
    month_quality_df = month_quality_df[month_quality_df["系统数量"] != 0]

    quality_df_len = len(quality_df)
    month_quality_df_len = len(month_quality_df)

    titles = ["品规(清洗后)", "求和项:数量", "", "品规(清洗后)", "系统数量", "商业数量", "差异", "备注", "",
              "年月+品规", "求和项:数量", "", "年月+品规", "系统数量", "商业数量", "差异", "备注"]
    data = pd.DataFrame("", index=range(max(quality_df_len, month_quality_df_len)), columns=titles)

    data.iloc[:quality_df_len, 0] = quality_df["品规(清洗后)"]
    data.iloc[:quality_df_len, 1] = quality_df["系统数量"]

    data.iloc[:quality_df_len, 3] = quality_df["品规(清洗后)"]
    data.iloc[:quality_df_len, 4] = quality_df["系统数量"]
    data.iloc[:quality_df_len, 5] = quality_df["商业数量"]
    data.iloc[:quality_df_len, 6] = quality_df["系统数量"] - quality_df["商业数量"]

    data.iloc[:month_quality_df_len, 9] = month_quality_df["年月+品规"]
    data.iloc[:month_quality_df_len, 10] = month_quality_df["系统数量"]

    data.iloc[:month_quality_df_len, 12] = month_quality_df["年月+品规"]
    data.iloc[:month_quality_df_len, 13] = month_quality_df["系统数量"]
    data.iloc[:month_quality_df_len, 14] = month_quality_df["商业数量"]
    data.iloc[:month_quality_df_len, 15] = month_quality_df["系统数量"] - month_quality_df["商业数量"]
    return data


def tidy_client_perspective(quality_df: pd.DataFrame, month_quality_df: pd.DataFrame):
    """整理透视表数据"""
    # 清理非中国中药的有效数据
    quality_df = quality_df[quality_df["商业数量"] != 0]
    month_quality_df = month_quality_df[month_quality_df["商业数量"] != 0]

    quality_df_len = len(quality_df)
    month_quality_df_len = len(month_quality_df)

    titles = ["品规(清洗后)", "求和项:数量", "", "品规(清洗后)", "商业数量", "系统数量", "差异", "备注", "",
              "年月+品规", "求和项:数量", "", "年月+品规", "商业数量", "系统数量", "差异", "备注"]
    data = pd.DataFrame("", index=range(max(quality_df_len, month_quality_df_len)), columns=titles)

    data.iloc[:quality_df_len, 0] = quality_df["品规(清洗后)"]
    data.iloc[:quality_df_len, 1] = quality_df["商业数量"]

    data.iloc[:quality_df_len, 3] = quality_df["品规(清洗后)"]
    data.iloc[:quality_df_len, 4] = quality_df["商业数量"]
    data.iloc[:quality_df_len, 5] = quality_df["系统数量"]
    data.iloc[:quality_df_len, 6] = quality_df["商业数量"] - quality_df["系统数量"]

    data.iloc[:month_quality_df_len, 9] = month_quality_df["年月+品规"]
    data.iloc[:month_quality_df_len, 10] = month_quality_df["商业数量"]

    data.iloc[:month_quality_df_len, 12] = month_quality_df["年月+品规"]
    data.iloc[:month_quality_df_len, 13] = month_quality_df["商业数量"]
    data.iloc[:month_quality_df_len, 14] = month_quality_df["系统数量"]
    data.iloc[:month_quality_df_len, 15] = month_quality_df["商业数量"] - month_quality_df["系统数量"]
    return data


def tidy_compare_result(quality_df: pd.DataFrame, path):
    """整理比对结果表的数据"""
    database_df = pd.read_excel(path, header=0)
    database_df = pd.merge(quality_df["品规(清洗后)"], database_df, on='品规(清洗后)', how='left')

    titles = ["商业名称", "商业责任人核查时间段（格式要求：20**年**月到20**年**月", "核查人", "流向收集人", "标准品规",
              "生产厂家", "营销系统数量", "收集数量", "差异数量（营销系统-收集）", "差异大类", "差异分类", "问题详述", "单价",
              "差异金额", "问题产生时间（格式要求：20**年**月到20**年**月）", "品种归属部门", "备注1", "备注2"]
    data = pd.DataFrame("", index=range(len(quality_df)), columns=titles)

    data["标准品规"] = quality_df["品规(清洗后)"]
    data["生产厂家"] = database_df["生产企业"]
    data["营销系统数量"] = quality_df["系统数量"]
    data["收集数量"] = quality_df["商业数量"]
    data["差异数量（营销系统-收集）"] = quality_df["系统数量"] - quality_df["商业数量"]

    data.loc[data["差异数量（营销系统-收集）"] > 0, "差异大类"] = "营销系统>收集"
    data.loc[data["差异数量（营销系统-收集）"] < 0, "差异大类"] = "营销系统<收集"
    data.loc[data["差异数量（营销系统-收集）"] == 0, "差异大类"] = "无差异"

    data.loc[data["差异数量（营销系统-收集）"] != 0, "问题产生时间（格式要求：20**年**月到20**年**月）"] = quality_df["年月"].apply(format_dates)

    data["单价"] = database_df["单价"]

    num_index = pd.to_numeric(data['单价'], errors='coerce').notna()
    data.loc[num_index, "差异金额"] = data.loc[num_index, "差异数量（营销系统-收集）"] * data.loc[num_index, "单价"]
    return data


def write_data(ws: Worksheet, data: pd.DataFrame):
    """将数据写入EXCEL表格"""
    titles = data.columns.tolist()
    if "color" in titles:
        titles.remove("color")
    for j, value in enumerate(titles):
        ws.write(0, j, value)
    for row_i, row_value in data.iterrows():
        row_i += 1
        color = STR2COLOR[row_value.pop("color") if "color" in row_value.index else ""]
        for col_i, value in enumerate(row_value):
            ws.write(row_i, col_i, value, color)


def format_dates(date_list: List):
    """格式化日期列表"""
    date_list.sort()
    start_date = datetime.datetime.strptime(date_list[0], '%Y-%m').date()
    end_date = datetime.datetime.strptime(date_list[-1], '%Y-%m').date()
    return f"{start_date.year}年{start_date.month}月到{end_date.year}年{end_date.month}月"


def main(path):
    zgzy_path = os.path.join(path, "中国中药表.xlsx")
    client_path = os.path.join(path, "客户表.xlsx")
    database_path = os.path.join(path, "中国中药产品库-湖南省.xlsx")
    print("读取中国中药数据表")
    zgzy_df, zgzy_df_tidy, zgzy_dict, client_database = get_zgzy_data(zgzy_path)
    print("读取终端客户的数据表")
    client_df, client_df_tidy, client_dict = get_client_data(client_path, client_database)
    print("开始比对数据，获取差异项")
    zgzy_different, client_different = compare_different(zgzy_dict, client_dict)
    print("整理中国中药及客户的数据")
    quality_df, month_quality_df = statistics_names(zgzy_df_tidy, client_df_tidy)
    print("定义核查报告表")
    wb = xlwt.Workbook()
    print("输出差异表")
    ws: Worksheet = wb.add_sheet("差异表")
    data = tidy_different_data(zgzy_df_tidy, zgzy_different, client_df_tidy, client_different)
    write_data(ws, data)
    print("输出差异细节表")
    ws: Worksheet = wb.add_sheet("差异细节表")
    data = tidy_different_detail(quality_df, month_quality_df)
    write_data(ws, data)
    print("保存核查报告")
    wb.save(os.path.join(path, "核查报告.xls"))
    print("定义结果表")
    wb = xlwt.Workbook()
    print("导入原始数据并标红")
    deal_excel(zgzy_df, wb.add_sheet("营销系统原始流向"), zgzy_different)
    deal_excel(client_df, wb.add_sheet("商业收集原始流向"), client_different)
    print("输出营销系统透视流向表")
    ws: Worksheet = wb.add_sheet("营销系统透视流向表")
    data = tidy_zgzy_perspective(quality_df, month_quality_df)
    write_data(ws, data)
    print("输出商业收集透视流向表")
    ws: Worksheet = wb.add_sheet("商业收集透视流向")
    data = tidy_client_perspective(quality_df, month_quality_df)
    write_data(ws, data)
    print("输出比对结果表")
    ws: Worksheet = wb.add_sheet("比对结果")
    data = tidy_compare_result(quality_df, database_path)
    write_data(ws, data)
    print("保存结果表")
    wb.save(os.path.join(path, "结果.xls"))
    print("程序已运行完毕")


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("-p", "--path", type=str, default=r"E:\NewFolder\liuxiang", help="数据所在的文件夹路径")
    opt = {key: value for key, value in vars(parser.parse_args()).items()}
    main(opt["path"])
