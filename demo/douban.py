# -*- coding: utf-8 -*-
'''
Created on 2017年11月18日

@author: Jeff Yang
'''
import requests
from lxml import etree
import xlwt
import time

workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('Douban', cell_overwrite_ok=True)
 
worksheet.write(0, 0, label='评论者')
worksheet.write(0, 1, label='时间')
worksheet.write(0, 2, label='内容')

header = {
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36',
        
               }
raw_cookies = 'bid=L_ruPCJrb6M; ll="118281"; _ga=GA1.2.1369563122.1493961363; __yadk_uid=WSNmWgdfVoljb1IWgxAwmPgBAjQCw0wW; gr_user_id=81385a4d-064c-4eff-9070-a091afc22d64; viewed="27061630_3351237"; ap=1; ps=y; ue="1786494804@qq.com"; dbcl2="139416577:58/z3yu8060"; ck=1mgP; _vwo_uuid_v2=C0E604A94F07F78D84A336DA3242F210|a20238c3f4c66252fe7b13a0f72d4133; _pk_ref.100001.4cf6=%5B%22%22%2C%22%22%2C1511002458%2C%22https%3A%2F%2Fwww.baidu.com%2Fs%3Fie%3DUTF-8%26wd%3D%25E8%25B1%2586%25E7%2593%25A3%22%5D; _pk_id.100001.4cf6=6ba701974904da67.1497420153.4.1511002486.1510989903.; _pk_ses.100001.4cf6=*; __utma=30149280.1369563122.1493961363.1510989897.1511002458.16; __utmb=30149280.0.10.1511002458; __utmc=30149280; __utmz=30149280.1510755557.13.10.utmcsr=baidu|utmccn=(organic)|utmcmd=organic; __utmv=30149280.13941; __utma=223695111.1369563122.1493961363.1510989897.1511002458.4; __utmb=223695111.0.10.1511002458; __utmc=223695111; __utmz=223695111.1510979642.2.2.utmcsr=baidu|utmccn=(organic)|utmcmd=organic|utmctr=%E8%B1%86%E7%93%A3; push_noty_num=0; push_doumail_num=0' 
cookies = {}    
for line in raw_cookies.split(';'):    
    key, value = line.split('=', 1)
    cookies[key] = value


init_url = "https://movie.douban.com/subject/2158490/comments?status=P"
count = int(etree.HTML(requests.get(init_url, headers=header, cookies=cookies).text).xpath('//div[@class="clearfix Comments-hd"]/ul/li[1]/span/text()')[0][3:-1])
# print(count)

start = 0
while count > start:
    url = "https://movie.douban.com/subject/2158490/comments?start=" + str(start) + "&limit=20&sort=new_score&status=P&percent_type="
    content = requests.get(url, headers=header, cookies=cookies).text
    html = etree.HTML(content)
    # print(html.xpath('//div[@class="mod-bd"]/div[1]/div[@class="comment"]/p/text()'))
    comments_list_lenth = len(html.xpath('//div[@class="mod-bd"]/div[@class="comment-item"]'))
    i = 1
    while i <= comments_list_lenth:
        name = html.xpath('//div[@class="mod-bd"]/div[' + str(i) + ']/div[@class="comment"]/h3/span[2]/a/text()')
        worksheet.write(i + start, 0, label=name)
        date = html.xpath('//div[@class="mod-bd"]/div[' + str(i) + ']/div[@class="comment"]/h3/span[2]/span[3]/text()')
        worksheet.write(i + start, 1, label=date)
        comment = html.xpath('//div[@class="mod-bd"]/div[' + str(i) + ']/div[@class="comment"]/p/text()')
        worksheet.write(i + start, 2, label=comment)
        i = i + 1
    start = start + 20
    print(start)
    time.sleep(2)
workbook.save('douban.xls')
print("Done!")
