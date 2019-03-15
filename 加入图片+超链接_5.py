# -*- coding: utf-8 -*-
"""
Created on Thu Feb  7 21:18:05 2019

@author: OGT428
"""
from bs4 import BeautifulSoup
import urllib.request
import random
import time
import xlwt


def url_maker(url,k):
    if k==0:
        return url
    else:
        url = 'https://movie.douban.com/top250?start='+str(k)+'&filter='
        return url
    
def run(url,line):
    def random_ip(url):
        req = urllib.request.urlopen(url)
        html_c = req.read()
        html_c= html_c.decode('utf-8')
        a = []
        soup = BeautifulSoup(html_c,"html.parser") 
        movieList=soup.find('div',attrs={'id':'list','style':'margin-top:15px;'})
        movie = movieList.find('tbody')
        mov = movie.find_all('td',attrs = {'data-title':'IP'})
        port = movie.find_all('td',attrs = {'data-title':'PORT'})
        for i in range(0,len(mov)): 
            c=str(mov[i].getText())+':'+str(port[i].getText())
            a.append(c)
        return c
        
    url_ip="https://www.kuaidaili.com/free/"    
    iplist = random_ip(url_ip)
    proxy_support = urllib.request.ProxyHandler({'http':random.choice(iplist)})
    opener = urllib.request.build_opener(proxy_support)
    opener.addheaders = [('User-Agent', 'Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.65 Safari/537.36')]
    urllib.request.install_opener(opener)
    response = urllib.request.urlopen(url)

    #info = []
    html_c = response.read()
    html_c= html_c.decode('utf-8')
    soup = BeautifulSoup(html_c,"html.parser") 
    m = soup.find('ol',attrs = {'class' : 'grid_view'})
   
    
    for movielist in m.find_all('li'):
        a = []
        j = 0
       # outputMode="{0:{3}^20}\t{1:{3}^20}\t{2:{3}<30}"
        mo=movielist.find("div",attrs = {'class':'hd'})
        moviename=mo.find("span",attrs = {'class':'title'}).getText()
        a.append(moviename)    
        wsheet.write(line,j,moviename)   #写入excel
        
        j = j+1
        
        score = movielist.find('span',attrs={'class':'rating_num'}).getText()
        a.append(score)
        
        wsheet.write(line,j,score)
        
        review = movielist.find('span',attrs = {"class":'inq'})
        if(review):
           
            temp = review.getText()
            a.append(temp)
            j = j+1
        else:
            temp ='None'
            a.append(temp)
            j = j+1
        wsheet.write(line,j,temp)
        
       # f = open(str(num)+'.jpg',"wb")
       # for href in movielist.find_all("a"):
        movielist_a = movielist.find("a")
        movielist_href = movielist_a.get('href')
       # print(acc)
        a.append(movielist_href)
        j = j+1
       # info = info + [a]
        wsheet.write(line,j,movielist_href)
        
#将图片路径信息过滤出来
        movielist_img = movielist.find('img')
        movielist_image = movielist_img.get('src')
       # print(movielist_image)
#二进制形式写入图片信息并保存
        f = open('E:/project1/movieimg/'+str(line)+'.jpg',"wb")
        img_code = urllib.request.urlopen(movielist_image)
        f.write(img_code.read())
        
       # url_img = 'E:/movieimg/'+str(line)+'.jpg'
       # image_data = .BytesIO(urllib.request.urlopen(url_img).read())
       # worksheet.insert_image('E'+str(line), url_img, {'image_data': image_data})
        line+=1
       # print(a)
       # print(info)
       # print(outputMode.format(a[0], a[1], a[2],chr(12288)))

#创建设置excel表格单元格格式
wbook = xlwt.Workbook('E:/project1/MOVIE.xls')
wsheet = wbook.add_sheet('TestSheet')
wsheet.col(0).width = 6000
wsheet.col(2).width = 20000
wsheet.col(3).width = 15000
wsheet.col(4).width = 10000
#tall_style = xlwt.easyxf('font:height 5000;')
#for img_line in range(250):
#first_row = wsheet.row(img_line)
#first_row.set_style(tall_style)
#wsheet.insert_image('E'+str(img_line), 'E:/movieimg/'+str(img_line)+'.jpg')
#worksheet2.insert_image('B20', r'c:\images\python.png')

#输入第一排标题数据
wsheet.write(0,0,'MOVIE_NAME')
wsheet.write(0,1,'SCORE')
wsheet.write(0,2,'REVIEW')
wsheet.write(0,3,'LINK')

#img_line = img_line+1
#for img_num in range(0,img_line):
#    book = xlsxwriter.Workbook('pict.xlsx')
#    sheet = book.add_worksheet('demo')
#    sheet.insert_image('E:/movieimg/'str(img_num)+'.jpg')

k=0
i=-24

url='https://movie.douban.com/top250'
while k<=225:
    url=url_maker(url,k)
    time.sleep(1)
    k+=25
    i+=25
    run(url,i)

#outputMode="{0:30}\t{1:20}\t{2:30}"
#print(outputMode.format('moviename', 'score', 'review',chr(12288)))

newexcle="E:/project1/MOVIE.xls"
wbook.save(newexcle)        
        