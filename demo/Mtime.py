# -*- coding: utf-8 -*-
'''
Created on 2017年11月18日

@author: Jeff Yang
'''
import requests
from lxml import etree
import xlwt

workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('Mtime', cell_overwrite_ok=True)
 
worksheet.write(0, 0, label='评论者')
worksheet.write(0, 1, label='内容')

page = 1
while page <= 10:
    if page == 1:
        url = 'http://movie.mtime.com/70233/reviews/short/hot.html'
    else:
        url = 'http://movie.mtime.com/70233/reviews/short/hot-' + str(page) + '.html'
    content = requests.get(url).text
    html = etree.HTML(content)
    comments_list_lenth = len(html.xpath('//dl[@id="tweetRegion"]/dd'))
    i = 1
    while i <= comments_list_lenth:
        name = html.xpath('//dl[@id="tweetRegion"]/dd[' + str(i) + ']/div/div[1]/div/p[1]/a/text()')
        print(name)
        worksheet.write(i + comments_list_lenth * (page - 1), 0, label=name)
        comment = html.xpath('//dl[@id="tweetRegion"]/dd[' + str(i) + ']/div/h3/text()')
        print(comment)
        worksheet.write(i + comments_list_lenth * (page - 1), 1, label=comment)
        i = i + 1
    page = page + 1
workbook.save('mtime.xls')
print("Done!")
