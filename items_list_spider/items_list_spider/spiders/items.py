from openpyxl import Workbook
from scrapy.http import FormRequest
import csv
import os.path
import scrapy
import glob
from decouple import config


class ItemsSpider(scrapy.Spider):
    name = 'items'
    allowed_domains = ['cs.deals']
    start_urls = [
        'https://cs.deals/ajax/botsinventory?appid=0'
    ]

    def start_requests(self):
        frmdata = {}
        url = "https://cs.deals/ajax/botsinventory?appid=0"
        headers = {
            'authority': 'cs.deals',
            'content-length': '0',
            'Content-Type': 'application/json',
            'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="98", "Google Chrome";v="98"',
            'accept': 'application/json, text/javascript, */*; q=0.01',
            'x-requested-with': 'XMLHttpRequest',
            'sec-ch-ua-mobile': '?0',
            'user-agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36',
            'sec-ch-ua-platform': '"Linux"',
            'origin': 'https://cs.deals',
            'sec-fetch-site': 'same-origin',
            'sec-fetch-mode': 'cors',
            'sec-fetch-dest': 'empty',
            'referer': 'https://cs.deals/trade-skins?fbclid=IwAR3U2drIIzXACaPD72nbE4ICVKX8y094aMoSSqRekDV_rN6AsToPXZXx-4M',
            'accept-language': 'en-US,en;q=0.9,vi;q=0.8',
            'cookie': config('COOKIE')
        }

        yield FormRequest(url, callback=self.parse, formdata=frmdata, headers=headers)

    def parse(self, response):
        items = response.json()["response"]["items"]
        for item in items[self.appId]:
            yield {
                "name": item["c"],
                "price": item["i"]
            }

    def close(self, reason):
        csv_file = max(glob.iglob('*csv'), key=os.path.getctime)

        wb = Workbook()
        ws = wb.active

        with open(csv_file, 'r') as f:
            for row in csv.reader(f):
                ws.append(row)

        wb.save(csv_file.replace('.csv', '') + '.xlsx')
