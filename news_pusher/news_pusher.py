import requests,config,json,unifined_login
session=unifined_login.auto_login()
#获取校内网的新闻（如教务处）
p=session.get("https://ehall.nbu.edu.cn/publicapp/sys/tzggapp/modules/ggll/cxlmxdggxx.do?LMDM=8AF61CA7C2D2A19EE0534B13160AB43C&pageSize=15&pageNumber=1")
if p.status_code==200:
    json_unehall=json.loads(p.text)
    if json_unehall['code']=='0':
        news_list_ehall=json_unehall['datas']['cxlmxdggxx']['rows']
        news_pagesize_ehall=json_unehall['datas']['cxlmxdggxx']['pageSize']
        news_pageNumber_ehall=json_unehall['datas']['cxlmxdggxx']['pageNumber']
        news_pageNumber_ehall=json_unehall['datas']['cxlmxdggxx']['totalSize']
    else:
        print('获取新闻失败')
        exit()
else:
    print('获取新闻失败')
    exit()
news_str=json.dumps(news_list_ehall)
fiobj=open('news.json','w')
fiobj.write(news_str)
fiobj.close()
print(news_list_ehall)