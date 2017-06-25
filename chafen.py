# encoding:utf-8
import requests
from bs4 import BeautifulSoup
import pandas
import json
from os import remove
import codecs

print('''
***************** GSTT *****************
* Grades System Of TIANYI Test. 
* By @Foldblade
* https://github.com/Foldblade/GSTT
****************************************
''')

f = open('setting.json', 'r')
setting = json.load(f)
f.close()
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
    'kaohao' : str(kaohao)}
    s = requests.session()
    response = s.post(url, data = postdata)
    response.encoding = 'utf-8'  # 修改编码为utf-8
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
for kaohao in all:
    try :
        get_total(setting["xm_name"],kaohao)
    except:
        pass


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
        data.sort(["总分"], ascending=False)
        data['排名'] = '=RANK(D:D,D:D)'
    writer = pandas.ExcelWriter('output.xlsx')
    data.to_excel(writer, 'Sheet1')
    writer.save()
    remove(result.csv)


print('Done!')

