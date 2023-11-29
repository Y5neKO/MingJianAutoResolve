import json

data = '''
{
    "status": "0",
    "t": "",
    "set_cache_time": "",
    "data": [
        {
            "ExtendedLocation": "",
            "OriginQuery": "114.114.114.114",
            "appinfo": "",
            "disp_type": 0,
            "fetchkey": "114.114.114.114",
            "location": "江苏省南京市 电信",
            "origip": "114.114.114.114",
            "origipquery": "114.114.114.114",
            "resourceid": "6006",
            "role_id": 0,
            "shareImage": 1,
            "showLikeShare": 1,
            "showlamp": "1",
            "titlecont": "IP地址查询",
            "tplt": "ip"
        }
    ]
}
'''

# 解析JSON数据
parsed_data = json.loads(data)

# 获取省和市信息
location = parsed_data["data"][0]["location"]
info, other = location.split(" ")[0:2]
province, city = info.split("省")[0:2]
