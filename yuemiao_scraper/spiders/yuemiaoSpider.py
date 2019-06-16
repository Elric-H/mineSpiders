from scrapy.spiders import Spider
from scrapy import Request
import json


class yuemiaoSpider(Spider):
    name = 'yuemiao'
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 '
                      'Safari/537.36 MicroMessenger/6.5.2.501 NetType/WIFI WindowsWechat QBCore/3.43.884.400 '
                      'QQBrowser/9.0.2524.400',
        'Accept': 'application/json, text/plain, */*',
        'st': '95a2bd3aaa0e5ab07c5455bd598a1000',
        'tk': '6042cea7abe50045ba5d14e21bd26489_552f1f236c40e7c20746f5c00dc9002d',
    }
    # vaccine.do疫苗详情查询接口返回数据:vaccineCode: code;departmentVaccineId: id
    # departmentWorkTimes2.do接种日期接口返回数据：subscirbeTime：id
    # department/detail.do 医院查询接口返回数据： depaCode：code
    url = 'https://wx.healthych.com/order/subscribe/add.do?vaccineCode=8803&vaccineIndex=1&linkmanId=1069828&subscribeDate=2019-05-23&subscirbeTime=891&departmentVaccineId=3181&depaCode=5101090088_daebd8c891c5c69d7767dbe01e5b813f'

    def start_requests(self):
        yield Request(self.url, headers=self.headers)

    def parse(self, response):
        datas = json.loads(response.body)
        print(datas, datas['ok'], datas['ok'] == False )
        if datas and datas['ok'] == False:
            yield Request(self.url, headers=self.headers, dont_filter=True)





