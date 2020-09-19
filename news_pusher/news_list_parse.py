import json,sqlite3,requests,unifined_login,os
def download_file(filename,url):
    print('正在下载:'+url)
    r = requests.get(url)
    with open(filename, 'wb') as f:
        f.write(r.content)
        f.close()
    print('下载完成:'+url)
def prase_newslist(newslist):
    for r in newslist:
        news_title=r['GGBT']
        news_ID=r['GGDM']
        news_DATE=r['FBSJ']
        news_LMID=r['LMDM']
        news_LMNAME=r['LMMC']
        news_FBBM=r['FBBM']
        if int(r['CONMENT_TYPE'])==2:
            news_url=r['GO_URL']
            print(news_url)
        else:
            news_url="https://ehall.nbu.edu.cn/new/indexnbu.html?type=5?ggdm="+news_ID
            news_content=prase_news(news_ID)
            print(news_content)
            #检测文章ID是否存在
            #下载附件
            for rs in news_content['fj']:
                if not os.path.exists('fj'):
                    os.mkdir('fj')
                rs_url="https://ehall.nbu.edu.cn/"+rs['fjurl']
                rs_name=rs['fjname']
                download_file('fj/'+rs_name,rs_url)

def prase_news(newsid):
    rep={"querySetting":[],"GGDM":newsid}
    reps=json.dumps(rep)
    url="https://ehall.nbu.edu.cn/publicapp/sys/tzggapp/ggll/loadNoticeDetailInfo.do?data="+reps
    session=unifined_login.cookies_login()
    p=session.get(url)
    if p.status_code==200:
        json_content=json.loads(p.text)
        if json_content['code']=='0':
            news_content=json_content['data']['GG_DATA']['GGNR']
            fj=[]
            if 'FJ_DATA' in json_content['data'].keys():
                print('有附件...')
                for r in json_content['data']['FJ_DATA']:
                    fj_obj={"fjid":r['id'],"fjname":r['name'],"fjdate":r['ts'],"fjurl":r['fileUrl']}
                    fj.append(fj_obj)
            return {"news_content":news_content,"fj":fj}
        else:
            print('解析新闻发生错误:'+json_content['msg'])
            return False
    else:
            print('解析新闻发生错误')
            return False
fiobj=open('news.json','r')
news_str=fiobj.read()
fiobj.close()
news_obj=json.loads(news_str)
prase_newslist(news_obj)



