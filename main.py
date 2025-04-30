import os
import time

import openpyxl
import pandas as pd
import numpy as np
import requests
import random
from lxml import etree
from openpyxl import Workbook
from tqdm import tqdm
def get_data(name,company,begin_year=2001,end_year=2021,sleep_time=random.random()+6):
    headers = {
        "Accept": "*/*",
        "Accept-Language": "en-US,en;q=0.9,zh-CN;q=0.8,zh;q=0.7,en-GB;q=0.6",
        "Connection": "keep-alive",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "DNT": "1",
        "Origin": "https://www.letpub.com.cn",
        "Referer": "https://www.letpub.com.cn/index.php?page=grant",
        "Sec-Fetch-Dest": "empty",
        "Sec-Fetch-Mode": "cors",
        "Sec-Fetch-Site": "same-origin",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36 Edg/126.0.0.0",
        "X-Requested-With": "XMLHttpRequest",
        "sec-ch-ua": "\"Not/A)Brand\";v=\"8\", \"Chromium\";v=\"126\", \"Microsoft Edge\";v=\"126\"",
        "sec-ch-ua-mobile": "?0",
        "sec-ch-ua-platform": "\"Windows\""
    }
    cookies = {
        "__utmz": "189275190.1707870561.1.1.utmcsr=bing|utmccn=(organic)|utmcmd=organic|utmctr=(not%20provided)",
        "PHPSESSID": "qtm70o5bckg67nc742sjmjjph7",
        "__utmc": "189275190",
        "_ga": "GA1.3.1737155656.1707870561",
        "_gid": "GA1.3.1103740081.1719458220",
        "_ga_MQPLXSWFB7": "GS1.1.1719458219.5.1.1719459491.0.0.0",
        "__utmt": "1",
        "__utma": "189275190.1737155656.1707870561.1719458219.1719458219.5",
        "__utmb": "189275190.2.10.1719458219",
        "_ga_JVDFHF8S2G": "GS1.3.1719465650.5.0.1719465650.0.0.0"
    }
    url = "https://www.letpub.com.cn/nsfcfund_search.php"
    params = {
        "mode": "advanced",
        "datakind": "list",
        "currentpage": "1"
    }
    data = {
        "page": "",
        "name": "",
        "person": name,
        "no": "",
        "company": company,
        "addcomment_s1": "",
        "addcomment_s2": "",
        "addcomment_s3": "",
        "addcomment_s4": "",
        "money1": "",
        "money2": "",
        "startTime": f"{begin_year}",
        "endTime": f"{end_year}",
        "province_main": "",
        "subcategory": "",
        "searchsubmit": "true"
    }
    response = requests.post(url, headers=headers, cookies=cookies, params=params, data=data)
    time.sleep(sleep_time)
    return response

def append_to_excel(name ,title, codes, project_type, company_names,file_name):
    wb = openpyxl.load_workbook(file_name)
    ws = wb.active

    # 将学科代码分组
    code_group = codes[:3]

    row = [name] + [code_group[0].split("：")[1]]+[title]+ [code_group[0].split("：")[1],code_group[1].split("：")[1],code_group[2].split("：")[1]] + [project_type] + [company_names]

    ws.append(row)

    # 保存Excel文件
    wb.save(file_name)
    # print(f"数据已写入 {file_name}")


if __name__ == '__main__':
    if not os.path.exists("output.xlsx"):
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        wb.save("output.xlsx")

    df = pd.read_excel(r"D:\桌面\2020名单.xlsx")
    time.sleep(500)
    for i in tqdm(range(len(df))):
        name = df.iloc[i]["姓名"]

        try:
            html_context = get_data(name,"").text
            tree = etree.HTML(html_context)
            titles = tree.xpath("//td[text()='题目']/following-sibling::td[1]/text()")
            primary_codes = tree.xpath("//td[text()='学科代码']/following-sibling::td[1]/text()")
            project_types = tree.xpath('//tr/td[5]/text()')
            company_names = tree.xpath(r'//tr[@style="background:#EFEFEF;"]/td[2]/text()')
            print(company_names )
            title_list = []
            code_list = []
            for title, primary_code in zip(titles, primary_codes):
                title_list.append(title)
                primary_code_parts = primary_code.split("，")
                for code in primary_code_parts:
                    code_list.append(code.strip())

            grouped_codes = [code_list[i:i + 3] for i in range(0, len(code_list), 3)]
            data_blocks = zip(titles, grouped_codes, project_types, company_names)
            for title, codes, project_type, company_name in data_blocks:
                append_to_excel(name,title, codes,project_type,  company_name,"output.xlsx")


        except Exception as e:
            print(f"出粗啦！人名{name}",)
            print("错误是：")
            print(e)


