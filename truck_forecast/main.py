from vo.Terminal import Terminal
from controller import predict_controller as pc

from flask import Flask, render_template, request
import json
import traceback
server = Flask(__name__)


@server.route('/get_bar', methods=['GET'])
def get_bar():
    # 下拉框
    selected_list = []
    for key, value in Terminal.terminal_base_info.items():
        selected_list.append({"code": key, "name": value[0], "select": False})
    selected_list[0]["select"] = True
    return render_template("terminal_info_bar.html", selected_list=selected_list)


@server.route('/get_bar_data', methods=['POST'])
def get_bar_data():
    try:
        terminal_code = request.form["terminal_code"]
        legend_type = request.form["legend_type"]
        # 获取柱状图数据
        if legend_type == "weight":
            data = {"legends": ["重箱", "空箱"]}
        elif legend_type == "eirtype":
            data = {"legends": ["出口", "进口"]}
        terminal = Terminal(terminal_code)
        data["terminal_name"] = terminal.name
        data.update(pc.get_data_for_bar(terminal_code, legend_type))
        return json.dumps({"success": True, "data": data}, ensure_ascii=False)
    except Exception as e:
        print(traceback.print_exc())
        return json.dumps({"success": False, "msg": str(e)}, ensure_ascii=False)


@server.route('/get_pie', methods=['GET'])
def get_pie():
    # 下拉框
    selected_list = []
    for key, value in Terminal.terminal_base_info.items():
        selected_list.append({"code": key, "name": value[0], "select": False})
    selected_list[0]["select"] = True
    result = {"selected_list": selected_list}
    # 柱状图的穿透链接
    params = request.args
    data_type = params.get("data_type")
    terminal_code = params.get("terminal_code")
    if data_type is not None and terminal_code is not None:
        result["terminal_code"] = terminal_code
        result["data_type"] = int(data_type)
    return render_template("terminal_info_pie.html", **result)


@server.route('/get_pie_data', methods=['POST'])
def get_pie_data():
    try:
        terminal_code = request.form["terminal_code"]
        data_type = int(request.form["data_type"])
        # 获取饼状图数据
        terminal = Terminal(terminal_code)
        data = {"terminal_name": terminal.name}
        data.update(pc.get_data_for_pie(terminal_code, data_type))
        return json.dumps({"success": True, "data": data}, ensure_ascii=False)
    except Exception as e:
        print(traceback.print_exc())
        return json.dumps({"success": False, "msg": str(e)}, ensure_ascii=False)


if __name__ == '__main__':
    # port可以指定端口，默认端口是5000
    # host写成0.0.0.0的话，其他人可以访问，代表监听多块网卡上面，默认是127.0.0.1
    server.run(debug=True, port=5000, host='0.0.0.0')
