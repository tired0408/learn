from dao.oracle_dao import oracle_dao
from vo.Terminal import Terminal

import math
import requests
import time
import random
detail_about_bar = None


def cal_distance(lng_a, lat_a, lng_b, lat_b):
    """
    根据经纬度计算两点间距离（只适用于两点都在北半球）
    :param lng_a: A点经度（百度地图）
    :param lat_a: A点纬度（百度地图）
    :param lng_b: B点经度（百度地图）
    :param lat_b: B点纬度（百度地图）
    :return: 单位km
    """
    r = 6371.004  # 地球半径KM
    c = math.sin(math.radians(lat_a))*math.sin(math.radians(lat_b))+math.cos(math.radians(lat_a)) \
        * math.cos(math.radians(lat_b))*math.cos(math.radians(lng_a-lng_b))
    distance = r*math.acos(c)
    return distance


def baidu_mutlti_route(origin_coordinate_list, detination_coordinate):
    """
    调用百度接口批量算路（坐标格式：[纬度，经度]）
    :param origin_coordinate_list: list 多个起点的位置，如：[[40.45, 116.34], [40.54, 116.35]]
    :param detination_coordinate: 单个目的地的坐标，如：[40.34, 116.45]
    :return:
    """
    # 模拟数据发布的时候要删除 TODO
    if True:
        a = random.random() * 5
        duration_list = [random.random()*a*3600 for i in range(len(origin_coordinate_list))]
        return duration_list
    orgin_str = '|'.join(','.join('%f' % oc for oc in ocl) for ocl in origin_coordinate_list)
    detination_str = ','.join('%f' % dc for dc in detination_coordinate)
    ak = 'zMUSpHlAay54daBTKbEb8pDe8chyLTpH'
    url = "http://api.map.baidu.com/routematrix/v2/driving?output=json&origins=%s&destinations=%s&ak=%s&tactics=12" \
          % (orgin_str, detination_str, ak)
    time.sleep(1)  # 停顿一秒
    rs = requests.get(url.replace("\n", "")).json()
    if rs["status"] != 0:
        raise Exception('调用百度API接口失败,错误信息：%s' % rs["message"])
    duration_list = [route["duration"]["value"]*1.2 for route in rs["result"]]
    return duration_list


def predict_homework(terminal_code):
    """
    根据码头代码，预估待进场的集装箱量情况
    :param terminal_code: 码头代码
    :return:
    """
    odb = oracle_dao()
    ocur = odb.cursor()
    terminal = Terminal(terminal_code)
    print("开始“%s”码头待进场集装箱作业量的预估。" % terminal.name)
    # 获取实时GPS表中有的车辆信息
    ocur.execute("""
            select * from tc_location where gpsdate > (sysdate-1/24)
            """)
    car_location_list = ocur.fetchall()
    if car_location_list is None:
        raise Exception("拖车GPS实时数据获取失败或无数据")
    car_no_list = [car[1] for car in car_location_list]
    # 获取EIR中的关于所有车辆的订单信息
    ocur_sql = """
            select * from (
            select tm.*,row_number() OVER(PARTITION BY tm.car_license ORDER BY tm.order_time desc) as row_flg from
            (
            select e.e_eirseqno,e.e_eirtype,e.e_status, 
            (case e.e_eirtype when 'E' then e.tc_in_dispatch_date when 'I' then e.tc_dispatch_date end) as order_time,
            (case e.e_eirtype when 'E' then e.e_indepotname_cn when 'I' then e.e_outdepotname_cn end) as terminal_name,
            (case e.e_eirtype when 'E' then e.e_indepot when 'I' then e.e_outdepot end) as terminal_code,
            (case e.e_eirtype when 'E' then e.in_driver when 'I' then e.driver end) as driver_name,
            (case e.e_eirtype when 'E' then e.e_in_truckno when 'I' then e.e_truckno end) as car_license,
            e.e_ves_vesselnamecn,e.e_ves_voyage,e.e_sizeofcntr,e.e_cntrno,e.e_two_eirno,e.e_ves_vesselnameen,
            (case e.e_eirtype when 'E' then e.e_in_emptyweightflag when 'I' then e.e_out_emptyweightflag end) 
            as emptyweightflag
            from eirinfo e where (e.e_eirtype='I' and e.mt_indate is null and e.mt_outdate is null) or (e.e_eirtype='E' 
            and e.e_status = 5 and e.e_sealno is not null and e.mt_indate is null and e.mt_outdate is null) 
            ) tm ) where row_flg = 1 and terminal_code = '%s' and order_time>sysdate-3
            """ % terminal.code
    ocur.execute(ocur_sql)
    data_content = []  # 存储处理后的订单信息
    while 1:
        el = ocur.fetchone()
        if el is None:
            break
        # 排除无近一小时GPS数据的车辆
        if el[7] not in car_no_list:
            continue
        # 获取GPS数据
        car_location = car_location_list[car_no_list.index(el[7])]
        # 排除距离码头直径50KM以外的车辆
        if cal_distance(float(car_location[4]), float(car_location[5]), terminal.bd_lng, terminal.bd_lat) > 50:
            continue
        # 排除当前速度为0的车辆（即可能未行车）
        if car_location[6] == 0:
            continue
        gps_timestamp = car_location[8].timestamp()
        data_content.append({"e_eirseqno": el[0], "e_eirtype": el[1], "gps_timestamp": gps_timestamp,
                             "e_ves_vesselname": el[8] if el[8] is not None else el[13],
                             "e_ves_voyage": el[9] if el[9] is not None else "未知",
                             "e_sizeofcntr": el[10], "car_license": el[7], "e_status": el[2], "e_cntrno": el[11],
                             "driver_name": el[6], "bd_lng": float(car_location[4]), "bd_lat": float(car_location[5]),
                             "e_two_eirno": el[12], "emptyweightflag": el[14] if el[14] is not None else "F"})
    if len(data_content) == 0:
        raise Exception("无'%s'码头的相关车辆的EIR信息。" % terminal.name)
    # 调用百度API批量算路
    print("开始进行批量算路")
    for i in range(math.ceil(len(data_content) / 50)):
        cll = data_content[i * 50:(i + 1) * 50]
        origin_coordinate_list = [[cl["bd_lat"], cl["bd_lng"]] for cl in cll]
        bd_route_time = baidu_mutlti_route(origin_coordinate_list, [terminal.bd_lat, terminal.bd_lng])
        for j in range(len(bd_route_time)):
            data_content[i * 50 + j]["book_entry_time"] = data_content[i * 50 + j]["gps_timestamp"] + bd_route_time[j]
    odb.close()
    return data_content


def get_data_for_bar(terminal_code, legend_type):
    """
    获取柱状图所需数据
    :param legend_type: 划分纬度{weight:空重箱，eirtype:进出口}
    :param terminal_code:
    :return:
    """
    global detail_about_bar
    detail_about_bar = predict_homework(terminal_code)  # 获取预估数据
    # 整理柱状图所需数据
    return_info = {}
    time_now_long = int(time.time())  # 获取当前时间戳（秒）
    x_value = []
    for i in range(6):
        x_value.append("%s-%s" % (time.strftime("%H:%M", time.localtime(time_now_long + i * 3600)),
                                  time.strftime("%H:%M", time.localtime(time_now_long + (i + 1) * 3600))))
    return_info["x_value"] = x_value
    return_info["y_value_list"] = [[0] * len(x_value), [0] * len(x_value)]  # [0]出口，[1]进口
    for i in range(len(detail_about_bar)):
        book_entry_time = detail_about_bar[i]["book_entry_time"]
        eirno_num = 1 if detail_about_bar[i]["e_two_eirno"] is None else 2
        time_interval = book_entry_time - time_now_long
        index = int(time_interval/3600)
        if index < 0:
            continue
        if legend_type == "eirtype":
            e_eirtype = detail_about_bar[i]["e_eirtype"]
            if "E" == e_eirtype:
                return_info["y_value_list"][0][index] += eirno_num
            elif "I" == e_eirtype:
                return_info["y_value_list"][1][index] += eirno_num
        elif legend_type == "weight":
            emptyweightflag = detail_about_bar[i]["emptyweightflag"]
            if "F" == emptyweightflag:
                return_info["y_value_list"][0][index] += eirno_num
            elif "E" == emptyweightflag:
                return_info["y_value_list"][1][index] += eirno_num
    print("结束“%s”码头待进场集装箱作业量的预估。" % terminal_code)
    return return_info


def get_data_for_pie(terminal_code, data_type):
    """
    获取饼状图所需数据
    :param data_type: 0全局1-6分别对应未来1-6小时
    :param terminal_code:
    :return:
    """
    if data_type == 0:
        data_content = predict_homework(terminal_code)  # 获取预估数据
    else:
        global detail_about_bar
        data_content = detail_about_bar
    # 整理成饼状图所需数据格式
    deal_data = {}
    vesselname_index = {}
    for i in range(len(data_content)):
        book_entry_time = data_content[i]["book_entry_time"]
        if data_type == 0:
            if book_entry_time > int(time.time()) + 3600*6:
                continue
        else:
            time_interval = book_entry_time - int(time.time())
            index = int(time_interval / 3600)
            if index < 0 or data_type != index+1:
                continue
        e_ves_vesselname = data_content[i]["e_ves_vesselname"]  # 船名
        e_ves_voyage = data_content[i]["e_ves_voyage"]  # 航次
        eirno_num = 1 if data_content[i]["e_two_eirno"] is None else 2
        if e_ves_voyage in deal_data:
            deal_data[e_ves_voyage] += eirno_num
        else:
            deal_data[e_ves_voyage] = eirno_num
            vesselname_index[e_ves_voyage] = [e_ves_vesselname, "出口" if data_content[i]["e_eirtype"] == "E" else "进口"]
    return_info = {"legend": [], "value": []}
    sort_list = sorted(deal_data.items(), key=lambda x: x[1], reverse=True)
    for sl in sort_list:
        return_info["legend"].append("%s/%s/%s" % (vesselname_index[sl[0]][0], sl[0], vesselname_index[sl[0]][1]))
        return_info["value"].append(sl[1])
    if data_type == 0:
        return_info["predict_time"] = "%s-%s" % (time.strftime("%H:%M", time.localtime(time.time())),
                                                 time.strftime("%H:%M", time.localtime(time.time()+3600*6)))
    else:
        return_info["predict_time"] = "%s-%s" % (time.strftime("%H:%M", time.localtime(time.time()+3600*(data_type-1))),
                                                 time.strftime("%H:%M", time.localtime(time.time()+3600*data_type)))
    print("结束“%s”码头待进场集装箱作业量的预估。" % terminal_code)
    return return_info
