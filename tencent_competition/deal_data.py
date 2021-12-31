import pandas as pd
import numpy as np
import time
import re # 正则表达式
import datetime

path = "E:/tencent_data/"
save_path = path + "deal_data/"
total_exposurel_log = "totalExposureLog.out"# 历史曝光日志数据文件6G
user_data = "user_data" # 用户特征属性文件3G
ad_static_feature = "ad_static_feature.out" # 广告静态数据
ad_operation = "ad_operation.dat" # 广告操作数据
test_sample = "test_sample.dat" # 测试数据
test_sample_b = "Btest_sample_new.dat"
def timestamp_to_str(timestamp):
    """
    时间戳转字符串
    :param timestamp:
    :return:
    """
    return time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(timestamp))
def timestr_to_timestamp(timestr):
    """
    字符串转时间戳
    :return:
    """
    return int(time.mktime(time.strptime(timestr, "%Y-%m-%d %H:%M:%S")))
def deal_ad_static_feature():
    """
    01对广告静态数据进行处理
    :return:
    """
    df = pd.read_csv(path+ad_static_feature, header=None, low_memory=False, sep="\t")
    old_row = df.shape[0]
    df = df.dropna(axis=0, how="any") # 去除有空值的行
    # 去除商品id为空的字段，即为-1，和出现多值的情况(有三条)
    df[3] = df[3].astype("str")
    df = df[~df[3].str.contains(",")]
    df[3] = df[3].astype("int64")
    df = df[~df[3].isin([-1])]
    # 去除广告行业id出现多值的情况
    df[5] = df[5].astype("str")
    df = df[~df[5].str.contains(",")]
    df[5] = df[5].astype("int64")
    # 去除创建时间异常的数据，异常基本为0
    df[1] = df[1].astype("int64")
    df = df[df[1]>100]
    new_row = df.shape[0]
    # 素材尺寸都只有一个，转化为int类型
    df[6] = df[6].astype("int64")
    # 输出为hdf文件
    df.to_hdf(save_path + ad_static_feature.replace(".out", ".h5"), mode="w", key="df")
    print("广告静态数据处理完成，原数据%s条，新数据%s条，删除了%s条" % (old_row,new_row,old_row-new_row))

def deal_ad_operation():
    """
    02对广告操作数据进行处理
    :return:
    """
    df = pd.read_csv(path+ad_operation, header=None, low_memory=False, sep="\t")
    old_row = df.shape[0]
    # 去除广告静态数据里面没有的广告
    ad_static_feature_data = pd.read_hdf(save_path + ad_static_feature.replace(".out", ".h5"),key="df")
    valid_ad_id = ad_static_feature_data[0]
    df = df[df[0].isin(valid_ad_id)]
    # 去除错误的时间日期，如20190230000000,并转化为时间戳
    def timestr_to_timestamp(time_str):
        if time_str == "0":
            return 0
        try:
            time_array = time.strptime(time_str, "%Y%m%d%H%M%S")  # 字符串转时间数组（time.struct_time）
            timestamp = int(time.mktime(time_array))  # 字符串转化为时间戳
        except ValueError as e:
            return -1
        return timestamp
    df[1] = df[1].map(lambda x:timestr_to_timestamp(str(x)))
    df[1] = df[1].astype("int64")
    df = df[~df[1].isin([-1])]
    # 填充创建时间为0的数据（从广告静态数据中获取）
    def fill_zero(data):
        if data[1] == 0:
            ad_id = data[0]
            data[1] = ad_static_feature_data[1][ad_static_feature_data[0] == ad_id]
        return data
    df = df.apply(lambda row:fill_zero(row),axis=1)
    # 去除只有修改，没有新建的广告数据
    lt = {}
    del_id = []
    for index, row in df.iterrows():
        if row[0] in lt.keys():
            lt[row[0]].add(row[2])
        else:
            lt[row[0]] = {row[2]}
    for obj in lt.items():
        if 2 not in obj[1]:
            del_id.append(obj[0])
    df = df[~df[0].isin(del_id)]

    new_row = df.shape[0]
    df.to_hdf(save_path + ad_operation.replace(".dat", ".h5"), mode="w", key="df")
    print("广告操作数据处理完成，原数据%s条，新数据%s条，删除了%s条" % (old_row, new_row, old_row - new_row))

def get_intersection():
    """
    03求广告静态数据及广告操作数据共有的广告数据
    :return:
    """
    df_1 = pd.read_hdf(save_path + ad_operation.replace(".dat", ".h5"), key="df")
    df_2 = pd.read_hdf(save_path + ad_static_feature.replace(".out", ".h5"), key="df")
    intersection_id = set(df_1[0]).intersection(set(df_2[0]))
    df_1 = df_1[df_1[0].isin(intersection_id)]
    df_2 = df_2[df_2[0].isin(intersection_id)]
    df_1.to_hdf(save_path + ad_operation.replace(".dat", ".h5"), mode="w", key="df")
    df_2.to_hdf(save_path + ad_static_feature.replace(".out", ".h5"), mode="w", key="df")
    print("处理完成！")

def deal_total_exposurel_log():
    """
    04对历史曝光数据进行处理
    :return:
    """
    reader = pd.read_csv(path+total_exposurel_log, header=None, low_memory=False, sep="\t", chunksize=10000000)
    i = 0
    for chunk in reader:
        old_row = chunk.shape[0]
        chunk = chunk.dropna(axis=0, how="any")  # 去除有空值的行
        # 去除广告静态数据里面没有的数据
        ad_static_feature_data = pd.read_hdf(save_path + ad_static_feature.replace(".out", ".h5"), key="df")
        valid_ad_id = ad_static_feature_data[0]
        chunk = chunk[chunk[4].isin(valid_ad_id)]
        i += 1
        chunk.to_hdf(save_path + total_exposurel_log.replace(".out", "_%s.h5" % i), mode="w", key="df")
        new_row = chunk.shape[0]
        print("已导出文件%s，原数据%s条，新数据%s条，删除了%s条" % (i,old_row,new_row,old_row-new_row))
def merge_exposurel_log():
    """
    05合并曝光日志文件,并进行去重处理
    :return:
    """
    # TODO 是否为空的判断有问题，takeeffect按道理不应该为空
    res = pd.DataFrame()
    for i in range(1,12):
        df = pd.read_hdf(save_path + total_exposurel_log.replace(".out", "_%s.h5" % i), key="df")
        res = res.append(df)

    res.drop_duplicates(keep="first", inplace=True)  # 去除完全重复的数据
    res.to_hdf(save_path + total_exposurel_log.replace(".out", ".h5"), mode="w", key="df")
    print("生成曝光日志文件，数据%s条" % res.shape[0])
def build_train_set():
    """
    06构建训练集
    :return:
    """
    op_df = pd.read_hdf(save_path + ad_operation.replace(".dat", ".h5"), key="df")
    op_df = op_df.loc[121:]
    st_df = pd.read_hdf(save_path + ad_static_feature.replace(".out", ".h5"), key="df")
    op_df.sort_values(by=[0,1], axis=0, ascending=True, inplace=True) # 对操作数据进行排序
    op_df.reset_index(drop=True, inplace=True) # 对索引进行重新排序
    columns = ["ad_id","create_time","material_size","industry_id","goods_type","good_id","ad_account_id",
               "target_time","target_people","bid","take_effect","end_effect","ad_state"]
    train_set = pd.DataFrame(columns=columns,dtype='object')
    invaild_id = []  # 记录有误的ID，修改在创建之前
    for index,row in op_df.iterrows():
        ad_id = row[0]
        if (not train_set.empty) and (row[0] == train_set.iloc[-1]["ad_id"]):
            last_index = train_set.iloc[-1].name
            if row[2] == 2:
                modify_field = row[3]
                if modify_field == 1:
                    train_set.loc[last_index,"ad_state"] = row[4]
                elif modify_field == 2:
                    if train_set.iloc[-1]["bid"].isnull():
                        raise Exception("新建的话，数据应该为空")
                    train_set.loc[last_index,"bid"] = row[4]
                elif modify_field == 3:
                    if train_set.iloc[-1]["target_people"].isnull():
                        raise Exception("新建的话，数据应该为空")
                    train_set.loc[last_index,"target_people"] = row[4]
                elif modify_field == 4:
                    if train_set.iloc[-1]["target_time"].isnull():
                        raise Exception("新建的话，数据应该为空")
                    train_set.loc[last_index,"target_time"] = row[4]
            else:
                #修改状态：
                #   修改为0：
                #       1、新建后遇到，填充结束时间，作为一条样本，并修改状态为0
                #       2、其他时候，最新状态是0，则不操作，
                #       3、填充结束时间，作为一个样本，并修改状态为0
                #   修改为1：
                #       1、新建后遇到，填充状态为1，作为一条样本
                #       2、其他时候，最新状态为0，新生成一个样本，开始时间为当前时间，清空结束时间，如果开始时间为空，则填充开始时间为当前时间
                #       状态不为0，填充状态为1，作为一条样本
                #修改其他：
                #   1、新建后遇到，填充结束时间，作为一条样本，同时生成新样本，开始时间为当前修改时间
                #   2、其他时候，遇到状态是0的，新生成一个样本，清空开始时间，清空结束时间
                #   遇到状态不是0的，填充结束时间，同时新生成一个样本，开始时间为修改时间
                modify_field = row[3]
                change_value = row[4]
                if modify_field == 1 : # 修改广告状态
                    change_value = int(change_value)
                    if change_value == 0:
                        if train_set.loc[last_index,"ad_state"] == 0:
                            continue
                        train_set.loc[last_index,"end_effect"] = row[1]
                        train_set.loc[last_index, "ad_state"] = change_value
                    else:
                        if train_set.loc[last_index,"ad_state"] == 1:
                            train_set.loc[last_index, "ad_state"] = change_value
                            continue
                        if train_set.loc[last_index,"take_effect"].isnull():
                            train_set.loc[last_index, "take_effect"] = row[1]
                        else:
                            train_set = train_set.append(train_set.iloc[-1],ignore_index=True)
                            last_index = last_index + 1
                            train_set.loc[last_index, "take_effect"] = row[1]
                            train_set.loc[last_index, "end_effect"] = None
                        train_set.loc[last_index, "ad_state"] = change_value
                else: # 修改其他
                    # TODO and后面是后加的
                    if train_set.loc[last_index, "ad_state"] == 0 and train_set.loc[last_index, "end_effect"] is not None:
                        train_set = train_set.append(train_set.iloc[-1], ignore_index=True)
                        last_index = last_index + 1
                        train_set.loc[last_index, "take_effect"] = None
                        train_set.loc[last_index, "end_effect"] = None
                    else:
                        train_set.loc[last_index, "end_effect"] = row[1]

                        train_set = train_set.append(train_set.iloc[-1], ignore_index=True)
                        last_index = last_index + 1
                        train_set.loc[last_index, "take_effect"] = row[1]
                        train_set.loc[last_index, "end_effect"] = None
                    # 更新修改信息
                    if modify_field == 2:
                        train_set.loc[last_index, "bid"] = int(change_value)
                    elif modify_field == 3:
                        train_set.loc[last_index, "target_people"] = change_value
                    elif modify_field == 4:
                        train_set.loc[last_index, "target_time"] = change_value
        else:
            modify_type = row[2]
            if modify_type != 2: # 广告一定是先创建再修改
                invaild_id.append(ad_id)
            create_time = row[1]
            ad_st_data = st_df.loc[st_df[st_df[0]==ad_id].index[-1]] # 广告对应的静态数据
            material_size = ad_st_data[6]
            industry_id = ad_st_data[5]
            goood_id = ad_st_data[3]
            goods_type = ad_st_data[4]
            ad_account_id = ad_st_data[2]
            target_time = None
            target_people = None
            bid = None
            take_effect = create_time
            end_effect = None
            ad_state = None  # 1正常，0失效
            # 填充数据
            modify_field = row[3]
            if modify_field == 1:
                ad_state = int(row[4])
            elif modify_field == 2:
                bid = int(row[4])
            elif modify_field == 3:
                target_people = row[4]
            elif modify_field == 4:
                target_time = row[4]
            row = pd.Series([ad_id,create_time,material_size,industry_id,goods_type,goods_id,ad_account_id,target_time,
                             target_people,bid,take_effect,end_effect,ad_state],index=columns,dtype="object")
            train_set = train_set.append(row, ignore_index=True)

        if index % 1000 == 0:
            print("已处理%d条。" % index)
    train_set.to_csv(save_path+"train_set.csv", sep="\t", header=None, index=None, encoding="utf-8")
    print("广告出现先修改后创建的情况，共有%s条" % len(invaild_id))

def get_target_time_str(value):
    """
    根据腾讯投放时间的规则，转化为可视化
    :return:
    """
    result = []
    for vl in value.split(","):
        time_list = []
        for sre_match in re.finditer(r'[1]+','{:048b}'.format(int(vl))):
            index =sre_match.span()
            start = 48 - index[1]
            end = 48 - index[0]
            time_list.append("%02d:%02d-%02d:%02d" % (start//2,start%2,end//2,end%2))
        result.append(",".join(time_list))
    return ";".join(result)

def get_submission():
    """
    TODO 进行预测，生成比赛所需文件
    :return:
    """
    df = pd.read_csv(path+test_sample_b, header=None, low_memory=False, sep="\t")
    res = pd.DataFrame()
    for index,row in df.iterrows():
        res = res.append(pd.Series([row[0], row[10]]), ignore_index=True)
    res[0] = res[0].astype("int64")
    res.to_csv(save_path+"submission.csv", sep=",", header=None, index=None, encoding="utf-8")
    print("已生成提交文件")


def normalized_data():
    """
    对测试集的数据进行处理
    :return:
    """
    df_adf = pd.read_hdf(save_path + "ad_static_feature.h5", key="df")
    df = pd.read_csv(path + "Btest_sample_new.dat", sep="\t", header=None, low_memory=False)
    df.drop([0, 1, 4, 7, 9], inplace=True, axis=1)
    # 对素材尺寸,出价进行归一化处理
    max_min_scaler = lambda x: (x - np.min(x)) / (np.max(x) - np.min(x))
    df[3] = df[[3]].apply(max_min_scaler)
    df[10] = df[[10]].apply(max_min_scaler)
    # 对投放时段、商品ID进行处理
    df["put_time"] = 0  # 投放时长（分）
    df["put_week"] = 0
    # df["has_gid"] = 0
    # df["without_gid"] = 0
    default_time = timestr_to_timestamp('2019-03-21 00:00:00')
    default_week = time.localtime(default_time).tm_wday
    for index, row in df.iterrows():
        # 商品ID分为有和没有两种
        # if row[6] == -1:
        #     df.loc[index, "without_gid"] = 1
        # else:
        #     df.loc[index, "has_gid"] = 1
        # 3月20日之前的广告预估3月21日,否则预估下一个自然日
        if row[2] < default_time:
            predict_week = default_week
        else:
            predict_week = time.localtime(row[2] + 60 * 60 * 24).tm_wday
        df.loc[index, "put_week"] = predict_week
        binary_str = '{:048b}'.format(int(row[8].split(",")[predict_week]))
        df.loc[index, "put_time"] = binary_str.count("1") * 48
        if (index + 1) % 1000 == 0:
            print("已处理%d条" % (index + 1))
    df.drop([2, 6, 8], inplace=True, axis=1)
    # 对投放时长和投放星期进行归一化处理
    df["put_time"] = df[["put_time"]].apply(max_min_scaler)
    df["put_week"] = df[["put_week"]].apply(max_min_scaler)
    # 处理商品类型，进行独热处理
    # 获取商品类型的映射函数
    # TODO 商品类型为1的数据，被清理了，但测试集里面有，后续可以找找原因
    prefix = "gt"
    goods_type_list = df_adf[4].value_counts().sort_index().index.tolist()
    goods_type_list_t = df[5].value_counts().sort_index().index.tolist()
    one_hot = pd.get_dummies(df[5], prefix=prefix, dtype="int64")
    # 添加训练集中有，测试集中没有的数据
    lt = set(goods_type_list).difference(goods_type_list_t)
    for l in lt:
        one_hot["%s_%d" % (prefix, l)] = 0
    # 将训练集没有商品类型归为ohter
    lt = set(goods_type_list_t).difference(set(goods_type_list))
    one_hot["%s_other" % prefix] = 0
    for l in lt:
        one_hot["%s_other" % prefix] = one_hot["%s_other" % prefix] + one_hot["%s_%d" % (prefix, l)]
    one_hot = one_hot.drop(["%s_%d" % (prefix, l) for l in lt], axis=1)
    df = df.drop(5, axis=1)
    df = df.join(one_hot)
    df.to_csv(save_path + test_sample_b.replace(".dat", ".csv"), sep="\t", index=None)
    print("测试集数据已处理完成。")

def deal_ad_id():
    """
    TODO 处理广告行业ID
    :return:
    """
    df_adf = pd.read_hdf(save_path + "ad_static_feature.h5", key="df")
    df = pd.read_csv(path + "Btest_sample_new.dat", sep="\t", header=None, low_memory=False)
    lt_1 = df_adf[5].value_counts()
    lt_1 = lt_1.sort_index()
    lt_2 = df[4].value_counts()
    lt_2 = lt_2.sort_index()
    print(lt_1.index.tolist())
    print(lt_2.index.tolist())

def deal_target_people():
    """
    TODO 处理定向人群
    :return:
    """
    df = pd.read_csv(path + "Btest_sample_new.dat", sep="\t", header=None, low_memory=False)
    map = {}
    for index,value in df[9].items():
        if value == "all":
            continue
        for v in value.split("|"):
            feature_name,feature_value = v.split(":")
            if feature_name in map:
                map[feature_name].update(feature_value.split(","))
            else:
                map[feature_name] = set(feature_value.split(","))
        if (index+1) % 1000 == 0:
            print("已处理%d条" % (index+1))
    for key,value in map.items():
        print(key, len(value))
    # print(df[9])

def generate_label():
    """
    根据构造好的训练集，生成标签
    :return:
    """
    start_time = time.time()
    # 获取样本
    train_df = pd.read_csv(save_path+"train_set.csv", sep="\t",low_memory=False)
    train_df = train_df[["ad_id", "take_effect", "end_effect","target_time"]]
    df = pd.read_hdf(save_path + "totalExposureLog.h5", key="df")
    df = df[[1,4,6,"time_str"]]
    # 对训练集进行处理
    # train_df.columns = ["ad_id", "create_time", "material_size", "industry_id", "goods_type", "good_id",
    #                     "ad_account_id","target_time", "target_people", "bid", "take_effect", "end_effect", "ad_state"]
    # train_df = train_df[~train_df["take_effect"].isnull()]  # 去除take_effect为空的行
    # 去除历史曝光日志最早时间之前的样本
    # default_stamp = timestr_to_timestamp("2019-02-16 00:00:00")
    # train_df = train_df[train_df["end_effect"]>default_stamp]
    # 判断投放时间小于24小时的，是否存在分隔两天,有24804条，暂时删除，TODO
    # train_df_copy = train_df[~train_df["end_effect"].isnull()].copy()
    # train_df_copy = train_df_copy[["ad_id","take_effect","end_effect"]]
    # train_df_copy["internal_time"] = train_df_copy["end_effect"] - train_df_copy["take_effect"]
    # train_df_copy = train_df_copy[train_df_copy["internal_time"]<24*3600]
    # train_df_copy["take_effect"] = train_df_copy[["take_effect"]].apply(lambda x: x - (x-time.timezone) % 86400)
    # train_df_copy["end_effect"] = train_df_copy[["end_effect"]].apply(lambda x: x - (x-time.timezone) % 86400)
    # train_df_copy = train_df_copy[train_df_copy["take_effect"] != train_df_copy["end_effect"]]
    # index_list = train_df_copy.index.tolist()
    # train_df.drop(index_list, axis=0, inplace=True)
    # i = 0
    # for index in index_list:
    #     row = train_df.loc[index].copy()
    #     train_df.drop(index, axis=0,inplace=True)
    #     take_effect = row["take_effect"]
    #     end_effect = row["end_effect"]
    #     new_row = row.copy()
    #     new_row["end_effect"] = take_effect - (take_effect - time.timezone) % 86400 + 86400-1
    #     train_df = train_df.append(new_row)
    #     new_row = row.copy()
    #     new_row["take_effect"] = end_effect - (end_effect - time.timezone) % 86400
    #     train_df = train_df.append(new_row)
    #     i = i + 1
    #     if i % 100 == 0:
    #         print("已处理%d条数据" % i)
    # train_df.to_csv(save_path+"train_set.csv", sep="\t", index=None, encoding="utf-8")
    # 提取标签
    train_df["put_time"] = 0
    train_df["put_week"] = 0
    train_df["label"] = None
    for index,row in train_df.iterrows():
        take_effect, end_effect, ad_id = row["take_effect"], row["end_effect"], row["ad_id"]
        if end_effect is None:  # 没有结束时间的样本，选取曝光数量最高的那一天
            total_data = df[df[4] == ad_id].copy()
            total_data[1] = total_data[[1]].apply(lambda x: x - (x-time.timezone) % 86400)
            label_series = total_data[1].value_counts()
            # 获取标签
            train_df.loc[index, "label"] = np.max(label_series)
            take_effect = label_series.idxmax()
            end_effect = take_effect + 24*3600 - 1
        elif end_effect - take_effect > 24 * 3600:
            continue
        else:
            total_data = df[df[4] == ad_id].copy()
            total_data = total_data[total_data[1]>take_effect]
            total_data = total_data[total_data[1]<end_effect]
            train_df.loc[index,"label"] = total_data.shape[0]

        put_week = time.localtime(take_effect).tm_wday
        train_df.loc[index, "put_week"] = put_week
        put_time = 0  # 投放时长(分)
        target_time = row["target_time"].split(",")[put_week]
        take_effect_zero = take_effect - (take_effect-time.timezone) % 86400
        for sre_match in re.finditer(r'[1]+', '{:048b}'.format(int(target_time))):
            start = 48 - sre_match.span()[1]
            end = 48 - sre_match.span()[0]
            start = take_effect_zero + start*30*60
            end = take_effect_zero + end*30*60
            start = max(start,take_effect)
            end = min(end,end_effect)
            if start > end:
                continue
            put_time = put_time + end - start
        train_df.loc[index, "put_time"] = put_time
        if (index+1) % 1000 == 0:
            end_time = time.time()
            print("已处理%d条数据。花费%d秒。" % (index+1,end_time-start_time))
    end_time = time.time()
    train_df.to_csv(save_path+"train_set_label.csv", sep="\t", index=None, encoding="utf-8")
    print("已处理完数据，耗时%d秒" % (start_time-end_time))


def normalized_data():
    """
    对训练集数据进行处理
    :return:
    """
    df_adf = pd.read_hdf(save_path + "ad_static_feature.h5", key="df")
    df = pd.read_csv(save_path+"train_set.csv", sep="\t",low_memory=False)
    df = df[["material_size","bid","goods_type","good_id"]]
    # 对商品id进行有和没有的区分，训练集全部具有商品ID，去掉该属性
    # df.loc[df[df["goods_type"] != -1].index.tolist(), "goods_type"] = 0
    # one_hot = pd.get_dummies(df["goods_type"], prefix="gid", dtype="int64")
    # df = df.join(one_hot)
    df = df.drop("goods_type", axis=1)
    # 对素材尺寸,出价进行归一化处理
    max_min_scaler = lambda x: (x - np.min(x)) / (np.max(x) - np.min(x))
    df["material_size"] = df[["material_size"]].apply(max_min_scaler)
    df["bid"] = df[["bid"]].apply(max_min_scaler)
    # 商品类型进行独热处理
    prefix = "gt"
    goods_type_list = df_adf[4].value_counts().sort_index().index.tolist()
    goods_type_list_t = df["good_id"].value_counts().sort_index().index.tolist()
    one_hot = pd.get_dummies(df["good_id"], prefix=prefix, dtype="int64")
    # 添加训练集中有，测试集中没有的数据
    lt = set(goods_type_list).difference(goods_type_list_t)
    for l in lt:
        one_hot["%s_%d" % (prefix, l)] = 0
    # 将训练集没有商品类型归为ohter
    lt = set(goods_type_list_t).difference(set(goods_type_list))
    one_hot["%s_other" % prefix] = 0
    for l in lt:
        one_hot["%s_other" % prefix] = one_hot["%s_other" % prefix] + one_hot["%s_%d" % (prefix, l)]
    one_hot = one_hot.drop(["%s_%d" % (prefix, l) for l in lt], axis=1)
    df = df.drop("good_id", axis=1)
    df = df.join(one_hot)
    df.to_csv(save_path + "train_set_other.csv", sep="\t", index=None, encoding="utf-8")

def merge_df():
    """
    合并训练集，生成最终版
    :return:
    """
    df_1 = pd.read_csv(save_path + "train_set_label.csv", sep="\t")
    df_1 = df_1[["put_time","put_week","label"]]
    df_2 = pd.read_csv(save_path + "train_set_other.csv", sep="\t")
    df = pd.concat([df_1,df_2],axis=1)
    df = df[~df["label"].isnull()]
    # 对投放时长和投放星期进行归一化处理
    max_min_scaler = lambda x: (x - np.min(x)) / (np.max(x) - np.min(x))
    df["put_time"] = df[["put_time"]].apply(max_min_scaler)
    df["put_week"] = df[["put_week"]].apply(max_min_scaler)
    df.to_csv(save_path + "train_set_finall.csv", sep="\t", index=None, encoding="utf-8")
# normalized_data()
# deal_ad_id()
# deal_target_people()
# df = pd.read_csv(save_path + ad_operation.replace(".dat", ".csv"), sep="\t",header=None)
# index = df[df[3]==4].index.tolist()
# df.iloc[index,4] = df.iloc[index,4].map(lambda x:get_target_time_str(x))
# deal_total_exposurel_log()
# get_intersection()
# df = pd.read_csv(path+test_sample, header=None, low_memory=False, sep="\t")
# merge_exposurel_log()
# df = pd.read_hdf(save_path + total_exposurel_log.replace(".out", ".h5"), key="df")
# get_submission()
# build_train_set()
# generate_label()
merge_df()
