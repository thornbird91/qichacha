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
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/57.0.2987.110 Safari/537.36',
        'Accept': 'text/html, application/xhtml+xml, application/xml; q=0.9, image/webp, image/apng, */*; q=0.8',
        'Accept-Encoding': 'gzip, deflate',
        'Connection': 'keep-alive',
        'Refer': 'http://www.qichacha.com/',
    }
    cook = {'Cookie': 'acw_tc=AQAAAGgrNSAcygoAkvi+PBmPDE22MQYR; UM_distinctid=15ffc6bab481ea-0a35f715b0cd86-396b4c0b-13c680-15ffc6bab49324; _uab_collina=151176844175827429607132; _umdata=2FB0BDB3C12E491D4C5F1DAFA82972B3C9EF950F96D22CD73891B83C6A9EE7F1149FA242ACBF79BFCD43AD3E795C914C4204B0C4F87E6E2661EC05B36556BA3C; PHPSESSID=nu478v7agulnuaspj9cm556pd7; hasShow=1; zg_did=%7B%22did%22%3A%20%2215ffc6bab7e144-0ad6f0e5baf8f-396b4c0b-13c680-15ffc6bab7f2ef%22%7D; CNZZDATA1254842228=2058476337-1511763115-%7C1511915470; zg_de1d1a35bfa24ce29bbf2c7eb17e6c4f=%7B%22sid%22%3A%201511919013101%2C%22updated%22%3A%201511919272595%2C%22info%22%3A%201511768435606%2C%22superProperty%22%3A%20%22%7B%7D%22%2C%22platform%22%3A%20%22%7B%7D%22%2C%22utm%22%3A%20%22%7B%7D%22%2C%22referrerDomain%22%3A%20%22www.qichacha.com%22%2C%22cuid%22%3A%20%2250912a6001bf14aed5cead7a5d52c876%22%7D'}
    req = requests.get(url=url, cookies=cook, headers=headers)
    html = req.content
    #url_cnt = BeautifulSoup(html, 'lxml')
    url_cnt = BeautifulSoup(html, 'html.parser')
    return url_cnt



def url_pro(dp_url, url_tail):
    txt1_wr = open('./tmp1.txt', 'a', encoding='gbk')
    dp = bs4_url(dp_url)
    #print(dp)

    ### get company name

    #search_list = dp.find_all('table', attrs={'class': 'm_srchList'})
    search_list = dp.table
    #search_list = dp.findAll('table')
    print(search_list)
    tbody = search_list[0].find_all('tbody')
    #print(tbody)
    tr_list = tbody[0].find_all('tr')
    print(len(tr_list))
    tr = tr_list[0]
    #print(tr)
    print(len(tr))
    dp_tit = tr.find_all('a', attrs={'class': 'ma_h1'})
    #dp_tit = dp.find_all('a', attrs={'class': 'ma_h1'})
    #print(dp_tit)
    if len(dp_tit):
        aa = str(dp_tit[0])
        #href = str(href_list[0])
        #href_list = re.findall('href="(.*?)"', aa)
        name = re.findall('<em><em>(.*?)</em></em>', aa)
        name0 = re.findall('blank">(.*?)<em>', aa)

        if len(name):
            if len(name0):
                name_ful = str(name0[0].strip()) + str(name[0].strip())
            else:
                name_ful = str(name[0])
        else:
            name_ful = ''


        if name_ful != url_tail:
            warn = '不匹配'
        else:
            warn = '匹配'

        ### get company status
        #sta_list_green  = tr.find('span', attrs={'class': 'ma_cbt_green m-l-xs'})
        #sta_list_orange = tr.find('span', attrs={'class': 'ma_cbt_orange m-l-xs'})
        sta_list  = re.findall('m-l-xs">(.*?)</span>', str(tr))
        print(sta_list)
        #if len(sta_list_green):
        #    sta = re.findall('>(.*?)</span>', str(sta_list_green[0]))
        #    print('aa')
        #    print(sta)
        #    print(sta_list_green[0])
        #elif len(sta_list_orange):
        #    sta = re.findall('>(.*?)</span>', str(sta_list_orange[0]))
        #    print('bb')
        #    print(sta)
        #    print(sta_list_orange[0])
        #else:
        #    sta = ['None']
        #print(sta)
        td_list = tr.find_all('td')
        #print(td_list)
        #print(len(td_list))
        #print(td_list[0])
        #print(td_list[1])
        #print(td_list[2])
        sta = re.findall('>(.*?)</span>', td_list[2])
        print(sta)
        print(url_tail, 'is OK')

    else:
        name_ful  = '查询不到'
        sta = 'NULL'
        warn = '不存在'
        print(url_tail, '不存在')

    #print(name_ful)

    wr_head = '输入公司名称' + '#' + '查询公司名称' + '#' + '经营状态' + '#' + '查询匹配度' + '\n'
    ele = url_tail + '#' + name_ful + '#' + str(sta[0]) + '#' + warn + '\n'

    txt1_wr.write(ele)
    txt1_wr.close()

    ### get company introduction
    #cmp_intr = dp.find_all('p', attrs={'class': 'm-t-xs'})
    #print(cmp_intr[0])



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
        row4 = line1[i].split('#')[3]
        #row3 = line2[i]
        sheet1.write(rows,0,row1)
        sheet1.write(rows,1,row2)
        sheet1.write(rows,2,row3)
        sheet1.write(rows,3,row4)
        rows +=1
    excel.save(exl_name)


def get_url(url_excel):
    url_head = 'https://www.qichacha.com/search?key='
    rd_xls = xlrd.open_workbook(url_excel)
    sheet = rd_xls.sheets()[0]
    nrows = sheet.nrows
    url_list = []
    cmp_name = []
    for i in range(1, nrows):
        #print(sheet.row_values(i))
        row_val = sheet.row_values(i)
        url_code = urllib.parse.quote(row_val[1])
        url = url_head + url_code
        url_list.append(url)
        cmp = row_val[1]
        cmp_name.append(cmp)
        #print(url)
    #print(url_list)
    #print(len(url_list))
    return (url_list, cmp_name)



url_head = 'https://www.qichacha.com/search?key='
url_tail = '杭州士兰微电子股份有限公司'
#url_tail = '杭州韵达贸易有限公司'
url_code = urllib.parse.quote(url_tail)
dp_url = url_head + url_code
url_pro(dp_url, url_tail)
#print(dp_url)



#if __name__ == '__main__':
#    rd_xls_name = 'addr1.xls'
#    wr_xls_name = 'result.xls'
#    (url_list, name_list) = get_url(rd_xls_name)
#    for i in range(len(url_list)):
#        #print(name_list[i])
#        url_pro(url_list[i], name_list[i])
#        time.sleep(2)
#    write_xls(wr_xls_name)



