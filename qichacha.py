import re
import urllib
import requests
import time

import xlrd
from bs4 import BeautifulSoup
import xlwt
from xlrd import open_workbook
from xlutils.copy import copy


def bs4_url(url):

    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.101 Safari/537.36',
        'Accept': 'text/html, application/xhtml+xml, application/xml; q=0.9, image/webp, image/apng, */*; q=0.8',
        'Accept-Encoding': 'gzip, deflate',
        'Connection': 'keep-alive',
    }
    req = requests.get(url=url, headers=headers)
    html = req.content
    url_cnt = BeautifulSoup(html, 'html.parser')
    return url_cnt



def url_pro(dp_url):
    txt1_wr = open('./tmp1.txt', 'a', encoding='gbk')
    dp = bs4_url(dp_url)
    dp_tit = dp.find_all('a', attrs={'class': 'ma_h1'})
    aa = str(dp_tit[0])
    href_list = re.findall('href="(.*?)"', aa)
    href = str(href_list[0])
    name = re.findall('<em><em>(.*?)</em></em>', aa)
    #print(href)
    print(str(name[0]))
    new_rtl = 'http://www.qichacha.com'+href
    print(new_rtl)
    sec_url = bs4_url(new_rtl)
    cmp_info = sec_url.find_all('td', attrs={'class': 'ma_left'})
    #print(cmp_info[13])
    st = str(cmp_info[13])
    sta = list(st.split(' '))
    #print(sta)
    while '' in sta:
        sta.remove('')
    print(sta[2])
    ana_sta = sta[2].strip()

    name_chn = str(name[0])
    sta_chn  = str(ana_sta)
    rtl_chn  = str(new_rtl)

    ele = str(name_chn) + '#' + rtl_chn + '#' + str(sta_chn) + '\n'
    txt1_wr.write(ele)
    time.sleep(1)



    txt1_wr.close()


def write_xls(exl_name):

    f_exl = open_workbook(exl_name, 'w+b')
    rows  = f_exl.sheets()[0].nrows
    excel = copy(f_exl)
    sheet1 = excel.get_sheet(0)

    #sheet1 = f_exl.add_sheet(u'sheet1', cell_overwrite_ok=True)
    txt1_rd = open('./tmp1.txt', 'r')
    #txt2_rd = open('./tmp2.txt', 'r')

    line1 = txt1_rd.readlines()
    #line2 = txt2_rd.readlines()
    for i in range(len(line1)):
        row1 = line1[i].split('#')[0]
        row2 = line1[i].split('#')[1]
        row3 = line1[i].split('#')[2]
        #row3 = line2[i]
        sheet1.write(rows,0,row1)
        sheet1.write(rows,1,row2)
        sheet1.write(rows,2,row3)
        rows +=1
    excel.save(exl_name)


def get_url(url_excel):
    #url_head = 'https://www.dianping.com/search/keyword/3/0_'
    url_head = 'https://www.qichacha.com/search?key='
    rd_xls = xlrd.open_workbook(url_excel)
    sheet = rd_xls.sheets()[0]
    nrows = sheet.nrows
    url_list = []
    for i in range(nrows):
        #print(sheet.row_values(i))
        row_val = sheet.row_values(i)
        url_code = urllib.parse.quote(row_val[0])
        url = url_head + url_code
        url_list.append(url)
        #print(url)
    #print(url_list)
    #print(len(url_list))
    return url_list


#url_head = 'https://www.qichacha.com/search?key='
#url_code = urllib.parse.quote('杭州士兰微电子股份有限公司')
#dp_url = url_head + url_code
#print(dp_url)


#html = urllib.request.urlopen(dp_url)
#print(html)
#print(dp_url)
#addr_chn = u'孩儿巷220号'
#addr_utf = addr_chn.encode('gbk')
#get_info(dp_url)


if __name__ == '__main__':
    rd_xls_name = 'addr.xls'
    wr_xls_name = 'result.xls'
    url_list = get_url(rd_xls_name)
    for i in range(len(url_list)):
        url_pro(url_list[i])
    write_xls(wr_xls_name)


#url_pro(dp_url)
#write_xls()
#get_url('addr.xls')

