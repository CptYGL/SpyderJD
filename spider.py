# -*- coding:utf-8 -*-
#@杨冠林 1316171511@qq.com
import io,sys,time,random
import requests                         #用于模拟网页请求,抓取
from openpyxl import load_workbook           #用于写入excel(why not csv???)
import lxml                             #html&xml解析库,方便处理数据
from bs4 import BeautifulSoup           #也是方便处理html页面(美味汤)
from json import loads                  #处理response-json转字典
#有乱码,网上查找得如下.需换输出格式
sys.stdout = io.TextIOWrapper(sys.stdout.buffer,encoding='utf-8')
merc_list = ['华为','OPPO','VIVO','小米','一加','苹果','黑鲨','三星','魅族','联想']
header = {'User-Agent': 'Mozilla/5.0'}
wb = load_workbook('data.xlsx')
wsheet = wb.worksheets[0]
wsheet.title = 'prefilt'
'''
取材来自-京东热卖 re.jd.com
取主流十个品牌->各取一页的爆款机型(16?)->各取质量最高用户前100条评价
csv结构:

'''

#分流器,进入对应的页面(搜索栏关键词方式)
def divider(merchan):
    url = ''
    if(merchan == merc_list[0]):
        url = 'https://search.jd.com/Search?keyword=华为手机'
    elif(merchan == merc_list[1]):
        url = 'https://search.jd.com/Search?keyword=OPPO手机'
    elif(merchan == merc_list[2]):
        url = 'https://search.jd.com/Search?keyword=VIVO手机'
    elif(merchan == merc_list[3]):
        url = 'https://search.jd.com/Search?keyword=小米手机'
    elif(merchan == merc_list[4]):
        url = 'https://search.jd.com/Search?keyword=一加手机'
    elif(merchan == merc_list[5]):
        url = 'https://search.jd.com/Search?keyword=苹果手机'
    elif(merchan == merc_list[6]):
        url = 'https://search.jd.com/Search?keyword=黑鲨手机'
    elif(merchan == merc_list[7]):
        url = 'https://search.jd.com/Search?keyword=三星手机'
    elif(merchan == merc_list[8]):
        url = 'https://search.jd.com/Search?keyword=魅族手机'
    elif(merchan == merc_list[9]):
        url = 'https://search.jd.com/Search?keyword=联想手机'
    else :
        url = None
        print('No Such Thing!!!\n')
    return url
#建立连接,取商品链接,返回集合,便于进入各个商品以读取所需信息
def get_info(the_url):
    response = requests.get(url=the_url,headers=header,verify=False)
    if(response.status_code == 200): print('Connection Established!\n')
    else:print('Connection failed!\n')
    response.encoding = response.apparent_encoding          #或者response.encoding = response.content.decode('utf-8')
    soup = BeautifulSoup(response.text,'lxml')
    #自营店waretype = 10    #J_goodsList > ul
    #参考https://www.cnblogs.com/yizhenfeng168/p/6979339.html
    goods = soup.select("li[ware-type='10']")
    for li in goods:
        prod_url = li.a.get('href')
        prod_id = prod_url.split('/')[3].split('.')[0]
        prod_price = li.i.text
        #print(prod_url,prod_price,prod_vol)
        get_comm(prod_id,prod_price)
#供get_info调用的子函数,真正读取详细评论
def get_comm(id,price):
    '''
    参考https://blog.csdn.net/weixin_42957905/article/details/106187180
    https://club.jd.com/comment/productPageComments.action?
    --- callback=fetchJSON_comment98&productId=10023108638660&
    --- score=0&sortType=5&page=0&pageSize=10&isShadowSku=0&rid=0&fold=1
    此为通过response获得的url,翻页时page=?会变,productid在不同产品时会变
    '''
    for page in range(10):
        time.sleep(random.randint(2,4))
        comm_url = 'https://club.jd.com/comment/productPageComments.action?callback=fetchJSON_comment98&'\
                'productId={_id}&score=0&sortType=5&page={_p}&pageSize=10&'\
                'isShadowSku=0&rid=0&fold=1'.format(_id=id,_p=page)      
        response = requests.get(url=comm_url,headers=header,verify=False).text
        #取字典/json(一页10个)
        prod_list = loads(response.lstrip('fetchJSON_comment98(').rstrip(');'))['productCommentSummary']
        comm_sum,comm_good,comm_mid,comm_bad = prod_list['commentCount'],prod_list['goodCount'],prod_list['generalCount'],prod_list['poorCount']
        comm_list = loads(response.lstrip('fetchJSON_comment98(').rstrip(');'))['comments']
        for com in comm_list:
            wsheet.append([id,com['referenceName'],price,comm_sum,comm_good,comm_mid,comm_bad,com['content']])
            print([id,com['referenceName'],price,comm_sum,comm_good,comm_mid,comm_bad,com['content']])
        wb.save('data.xlsx')
    

#主程序,调用即可
while(1):
    url = divider(input(str(merc_list)+'\n你要哪个牌子的?:'))
    if url == None:
        break
    else:
        get_info(url)
