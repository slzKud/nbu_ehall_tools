import requests,config,json,unifined_login,news_list_parse,sys,os
if len(sys.argv)>=2:
    if sys.argv[1]=="first" or not os.path.exists('config_user.json'):
        print('设置为初次同步模式')
        config.content_flag=False
        config.fj_flag=False
session=unifined_login.auto_login()
#获取校内网的新闻（如教务处）
for i in range(1,config.page_deepth+1):
    p=session.get("https://ehall.nbu.edu.cn/publicapp/sys/tzggapp/modules/ggll/cxlmxdggxx.do?LMDM=8AF61CA7C2D2A19EE0534B13160AB43C&pageSize={1}&pageNumber={0}".format(i,config.page_size))
    if p.status_code==200:
        json_unehall=json.loads(p.text)
        if json_unehall['code']=='0':
            news_list_ehall=json_unehall['datas']['cxlmxdggxx']['rows']
            news_pagesize_ehall=json_unehall['datas']['cxlmxdggxx']['pageSize']
            news_pageNumber_ehall=json_unehall['datas']['cxlmxdggxx']['pageNumber']
            news_pageNumber_ehall=json_unehall['datas']['cxlmxdggxx']['totalSize']
            print('获取新闻列表成功，开始抓取列表...')
            news_list_parse.prase_newslist(news_list_ehall)
            
        else:
            print('获取新闻失败')
            exit()
    else:
        print('获取新闻失败')
        exit()
