# encoding:utf-8
import requests
from bs4 import BeautifulSoup
import pandas
import json
from os import remove
import codecs
import time
import random

print('''
***************** GSTT *****************
* Grades System Of TIANYI Test. 
* By @Foldblade
* https://github.com/Foldblade/GSTT
****************************************
''')

start_time = time.clock()
user_agaents = ['Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/47.0.2526.73 Safari/537.36',
                'Mozilla/5.0 (Windows NT 6.3; rv:36.0) Gecko/20100101 Firefox/36.04',
                'Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US) AppleWebKit/533.20.25 (KHTML, like Gecko) Version/5.0.4 Safari/533.20.27',
                'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10; rv:33.0) Gecko/20100101 Firefox/33.0',
                'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2227.1 Safari/537.36',
                'Mozilla/5.0 (X11; CrOS i686 3912.101.0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/27.0.1453.116 Safari/537.36',
                'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2227.0 Safari/537.36',
                'Mozilla/5.0 (X11; Linux i586; rv:31.0) Gecko/20100101 Firefox/31.0',
                'Opera/9.80 (X11; Linux i686; Ubuntu/14.10) Presto/2.12.388 Version/12.16',
                'Opera/9.80 (Windows NT 6.0) Presto/2.12.388 Version/12.14']

try :
    f = open('setting.json', 'r')
    setting = json.load(f)
    f.close()
except :
    print('Bad json.\nDid you read \'ReadMe.md\' carefully ???\n你们这样是不行的啊\nI\'m angry !')
    time.sleep(30)

# 读取设置
wen_ke = setting['wen_ke']
# print(wen_ke)
li_ke = setting['li_ke']
# print(li_ke)

if setting["division"] == 1:
    f = codecs.open('wen.csv', 'w+', 'utf-8')
    f.write('考号' + ',' + '姓名' + ',' + '班级' + ',' + '总分' + '\n')
    f.close()
    f = codecs.open('li.csv', 'w+', 'utf-8')
    f.write('考号' + ',' + '姓名' + ',' + '班级' + ',' + '总分' + '\n')
    f.close()
else:
    f = codecs.open('result.csv', 'w+', 'utf-8')
    f.write('考号' + ',' + '姓名' + ',' + '班级' + ',' + '总分' + '\n')
    f.close()
# 清空历史记录

def get_total(xm_name,kaohao):
    url = 'http://tya.lhsvr.cn/index.php'
    postdata = {
    'xm_name' : str(xm_name),
    'kaohao' : str(kaohao),
    "headers": {
            "User-Agent": random.choice(user_agaents)}
    } # 好玩起见加了一段对user_anaent的切换
    s = requests.session()
    try :
        response = s.post(url, data=postdata, timeout=10)
        response.encoding = 'utf-8'  # 修改编码为utf-8
    except :
        print('Time out.Please check your connection to the Internet.')
        print('Exit in 10s.')
        time.sleep(10) # 网络中断，10s后自动退出
        exit(1)
    r = response.text
    webdata = BeautifulSoup(r, 'html.parser')
    #print('Data get.Analyzing...')
    find_td = webdata.find_all('td')
    #print('Detail get.Analyzing...')
    # print(webdata)
    # print(find_td)
    name = str(find_td[3].get_text())  # 姓名寻得
    # print(name)
    # school = str(find_td[5].get_text())  # 学校寻得
    total = float(find_td[7].get_text()) # 总分寻得
    # print(total)
    class_num = str(kaohao)[6:8]
    # print(class_num)
    if setting["division"] == 1:
        if int(class_num) in wen_ke :
            f = codecs.open('wen.csv', 'a+', 'utf-8')
            f.write(str(kaohao) + ',' + str(name) + ',' + str(class_num) + ',' + str(total) + '\n')
            f.close()
        else :
            f = codecs.open('li.csv', 'a+', 'utf-8')
            f.write(str(kaohao) + ',' + str(name) + ',' + str(class_num) + ',' + str(total) + '\n')
            f.close()
    else:
        f = codecs.open('result.csv', 'a+', 'utf-8')
        f.write(str(kaohao) + ',' + str(name) + ',' + str(class_num) + ',' + str(total) + '\n')
        f.close()
    # 记录为csv
    print(str(class_num) + '-' + str(name) + ' saved.' )


all = range(setting["kaohao"][0],setting["kaohao"][1])
all_num = float(setting["kaohao"][1] - setting["kaohao"][0])
count = 0
for kaohao in all:
    try :
        get_total(setting["xm_name"],kaohao)
    except:
        print('Blank or Bad data. Jump.')
    count = count + 1
    percentage = count / all_num
    print('[' + '=' * int(percentage * 30) + '>' + '-' * int(30 - int(percentage * 30)) + ']' + str(int(percentage * 100)) + '%')


if setting["division"] == 1:
    df1 = pandas.read_csv('wen.csv')
    df2 = pandas.read_csv('li.csv')
    if setting["rank"] == 1:
        df1 = df1.sort_values("总分", ascending=False)
        df1['排名'] = '=RANK(D:D,D:D)'  # 添加排名列、填写函数求排名
        df2 = df2.sort_values("总分", ascending=False)
        df2['排名'] = '=RANK(D:D,D:D)'
    writer = pandas.ExcelWriter('output.xlsx')
    df1.to_excel(writer, '文科', index=False)
    df2.to_excel(writer, '理科', index=False)
    writer.save()
    remove('wen.csv')
    remove('li.csv')
else :
    data = pandas.read_csv('result.csv')
    if setting["rank"] == 1:
        data = data.sort_values("总分", ascending=False)
        data['排名'] = '=RANK(D:D,D:D)'
    writer = pandas.ExcelWriter('output.xlsx')
    data.to_excel(writer, 'Sheet1', index=False)
    writer.save()
    remove('result.csv')


print('Done!'+ time.ctime())
end_time = time.clock()
print('Running time :' + str(int(end_time - start_time)) + 's .')
print('Exit in 30 s.')
time.sleep(30)
