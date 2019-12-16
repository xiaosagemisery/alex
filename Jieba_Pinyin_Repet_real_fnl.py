import jieba
from xpinyin import *
from datetime import datetime
import os
from bs4 import BeautifulSoup
import requests
import lxml
import pyautogui

# coding = utf-8

address=[]
r = requests.get(url = "https://www.medtronic.com/cn-zh/your-health/conditions.html")#一级界面链接
soup = BeautifulSoup(r.text,'lxml')

with open(r'C:\win10input\1.txt','a',encoding='utf-8') as fe:
    for i in (soup.find_all('a',href=re.compile("/cn-zh/your-health/conditions"))):#用正则搜索二级链接
        if i.get('href') =="/cn-zh/your-health/conditions.html":
            continue
        else:
            address.append("https://medtronic.com"+i.get('href'))
            for link in address:
                r1=requests.get(url=link)
                soup1 = BeautifulSoup(r1.text,'lxml')
                i1=(soup1.find_all('div','pad vert-space'))
                m=len(i1)
                for n in range(m-1):
                    fe.write(i1[n].text)

##爬取所有二级页面上的文字

punctuation = '\”“。、！，：；？"\‘,.1234567890qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM'

def removepunctuation(txt):
    txt = re.sub(r'[{}]+'.format(punctuation), '', txt)
    return txt.strip().lower()

p = Pinyin()
#定义要去除的元素

bblist = []
ddlist = []
gglist = []

with open('C:/win10input/1.txt',mode = 'r+',encoding='utf-8') as text4:
    for line in text4.readlines():
        line = removepunctuation(line)
        d = jieba.cut(line.strip(), cut_all=False)
        aa = " ".join(d)
        aalist = aa.split(" ")
        bblist = bblist+aalist
        cclist = list(set(bblist))
    for i in cclist:
        if len(i)<2:
            continue
        else:
            ddlist.append(i)
#jieba分词，再去除标点和数字以及长度只有1的词

os.chdir("C:/win10input/output/")
with open('MDTInput'+datetime.now().strftime('%Y%m%d %H%M%S')+'.txt',"ab+") as f:
    for i in ddlist:
        e = p.get_pinyin(i, "'")
        #print (e.capitalize()+"  "+i)
        f.write("'".encode('utf-8')+e.encode('utf-8')+" ".encode('utf-8')+
                i.encode('utf-8')+"\r\n".encode('utf-8'))
#给每个词加上拼音

file_dict = {}
file_dir = 'C:/win10input/output'
lists = os.listdir(file_dir)
for i in lists:
    ctime = os.stat(os.path.join(file_dir, i)).st_ctime
    file_dict[ctime] = i
max_ctime = max(file_dict.keys())
lastest_file=(file_dict[max_ctime])
print(lastest_file)
#定义最新创建的文件的文件名

os.chdir("C:/win10input/output/")
with open('sougouinternaldict.txt') as text5,open(lastest_file,mode='r+',encoding="utf-8") as text6:
    eelist=text5.readlines()
    hhlist=text6.readlines()
    gglist=list((set(hhlist).difference(set(eelist))))
    text6.seek(0)
    text6.truncate()
    for i in gglist:
        text6.write(i)
#和搜狗基本词库进行对比，只留下和搜狗词库中不同的词

os.startfile(r'C:\Users\sheny32\Downloads\Release_V2.6_Windows\深蓝词库转换.exe')
#自动打开词库转换软件

