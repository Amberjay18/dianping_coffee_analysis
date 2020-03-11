# -*- coding: utf-8 -*-
from openpyxl import load_workbook
import requests
import json

wb = load_workbook('your_file_name.xlsx')
sheet = wb.active
# 在第三列前面插入2列
sheet.insert_cols(3,2)
sheet['C1'] = 'location'
sheet['D1'] = 'area'

for geo in data['address']:
    parameters = {
        'address': geo,
        'city': '成都市',
        'key': your key
    }
    url = 'https://restapi.amap.com/v3/geocode/geo'
    #对高德地图web服务的api进行请求
    try:
        res = requests.get(url,params=parameters)
        #获取json格式的返回结果
        results = json.loads(res.text) 
        #返回经纬度的是location这个参数
        location = results['geocodes'][0]['location']
        #longitude = location.split(',')[0]
        #latitude = location.split(',')[1]
        area = results['geocodes'][0]['district']
        sheet.append([location,area])
                  
    except IndexError:
        sheet.append(['-','-'])
        continue
    
wb.save('your_file_name.xlsx')

