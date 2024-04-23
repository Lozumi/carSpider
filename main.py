import json
import requests
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.workbook import Workbook
from parsel import Selector
import pandas as pd
import time

# 搜索汽车名称url
get_car_id_url = "https://www.dongchedi.com/search?keyword={car_name}&currTab=1&city_name={city_name}&search_mode=history"
# 汽车详情url
get_car_detail_url = "https://www.dongchedi.com/motor/pc/car/series/car_list?series_id={car_id}&city_name={city_name}"
# 车友评论url
get_carfrind_comment = "https://www.dongchedi.com/motor/pc/car/series/get_review_list?series_id={car_id}&sort_by=default&only_owner=0&page={page}&count={count}"

# headers必须要有
headers = {
    'pragma': 'no-cache',
    'accept-language': 'zh-CN,zh;q=0.9',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.142 Safari/537.36',
    'accept': '*/*',
    'cache-control': 'no-cache',
    'authority': 'www.dongchedi.com',
    'referer': 'https://www.dongchedi.com/auto/series/3736',
}


# 获取车辆id的函数，最后返回id
def get_car_id(car_name, city_name):
    carid_url = get_car_id_url.format(car_name=car_name, city_name=city_name)
    response = requests.get(url=carid_url, headers=headers).text
    selector = Selector(text=response)
    car_message = selector.css('''.dcd-car-series a::attr(data-log-click)''').get()
    car_message = json.loads(car_message)
    car_id = car_message.get("car_series_id")
    return car_id


# 获取车友评论
def get_car_frind_comment(car_id, total_pages=30, count=50):
    car_frind_list = []
    for page in range(1, total_pages + 1):
        carfrind_url = get_carfrind_comment.format(car_id=car_id, page=page, count=count)
        response = requests.get(url=carfrind_url, headers=headers)
        if response.status_code == 200:
            try:
                response_json = response.json()
                car_frind_comment_list = response_json.get("data", {}).get("review_list", [])
                # 检查评论列表是否为空，如果为空，则跳过本次请求，继续下一次请求
                if not car_frind_comment_list:
                    print("未获取到车友评论信息，跳过本次请求。")
                    continue
                for car_frind_comment in car_frind_comment_list:
                    car_frind_dict = {}
                    buy_car_info = car_frind_comment.get("buy_car_info")
                    if buy_car_info:
                        bought_time = buy_car_info.get("bought_time", "")
                        location = buy_car_info.get("location", "")
                        price = buy_car_info.get("price", "")
                        series_name = buy_car_info.get("series_name", "")
                        car_name = buy_car_info.get("car_name", "")
                        car_frind_dict["成交时间"] = bought_time
                        car_frind_dict["地点"] = location
                        car_frind_dict["价格"] = price
                        car_frind_dict["系列名称"] = series_name
                        car_frind_dict["车型名称"] = car_name
                        car_content = car_frind_comment.get("content", "")
                        car_frind_dict["车主评论"] = car_content
                        car_frind_list.append(car_frind_dict)
            except json.JSONDecodeError:
                print("JSON解析错误")
        else:
            print("请求失败:", response.status_code)
        # 添加延迟，避免请求被关闭
        time.sleep(1)  # 在每次请求之间添加2秒的延迟
    return car_frind_list


# 获取车辆详情
def get_car_detail(car_id, city_name):
    car_detail = get_car_detail_url.format(car_id=car_id, city_name=city_name)
    response = requests.get(url=car_detail, headers=headers).json()
    online_all_list = response.get("data").get("tab_list")[0].get("data")
    car_type_list = []
    for car_cls in online_all_list:
        car_type_dict = {}
        car_cls = car_cls.get("info")
        if car_cls.get("id"):
            car_name = car_cls.get("series_name")
            car_type = car_cls.get("car_name")
            price = car_cls.get("price")
            owner_price = car_cls.get("owner_price")
            dealer_price = car_cls.get("dealer_price")
            upgrade = car_cls.get("upgrade_text")
            tags = "".join(car_cls.get("tags"))
            configure_list = car_cls.get("diff_config_with_no_pic")
            if configure_list is None:
                configure = ""
            else:
                configure = ", ".join([i.get('config_group_key') + "-" + i.get('config_key') for i in configure_list])
            car_type_dict["车辆名称"] = car_name
            car_type_dict["车辆类型"] = car_type
            car_type_dict["官方指导价"] = price
            car_type_dict["经销商报价"] = dealer_price
            car_type_dict["车主参考价"] = owner_price
            car_type_dict["车辆升级类型"] = upgrade
            car_type_dict["车辆标签"] = tags
            car_type_dict["车辆配置"] = configure
            car_type_list.append(car_type_dict)
    return car_type_list


# 保存为json文件
def save_json(car_name, text):
    json_text = json.dumps(text, ensure_ascii=False)
    with open(car_name + ".json", "w", encoding="utf-8") as f:
        f.write(json_text)
        print(car_name + " JSON文件保存成功")


# 保存为excel表格
def save_excel(car_name, carinfo):
    wb = Workbook()
    ws_detail = wb.active
    ws_detail.title = "车辆详细信息"
    df_detail = pd.DataFrame(carinfo["车辆详细信息"])
    write_to_sheet(ws_detail, df_detail)

    ws_frind = wb.create_sheet(title="车主成交信息")
    df_frind = pd.DataFrame(carinfo["车主成交信息"])
    write_to_sheet(ws_frind, df_frind)

    wb.save(car_name + ".xlsx")
    print(car_name + " Excel文件保存成功")


def write_to_sheet(ws, df):
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)


# 启动函数
def main(car_name, city_name, export_format="json"):
    car_id = get_car_id(car_name=car_name, city_name=city_name)
    carinfo = {
        "车辆详细信息": get_car_detail(car_id=car_id, city_name=city_name),
        "车主成交信息": get_car_frind_comment(car_id=car_id)
    }
    # if export_format == "json":
    #     save_json(car_name, carinfo)
    # elif export_format == "excel":
    #     save_excel(car_name, carinfo)
    # else:
    #     print("导出格式错误，请选择json或excel。")
    save_json(car_name, carinfo)
    save_excel(car_name, carinfo)


if __name__ == '__main__':
    car_list = ["宋PLUS DM-i"]
    for i in car_list:
        main(i, "西安", export_format="excel")
