import json,requests,unifined_login,os,db_connect,config,push,time
def download_file(filename,url):
    print('正在下载:'+url)
    r = requests.get(url)
    with open(filename, 'wb') as f:
        f.write(r.content)
        f.close()
    print('下载完成:'+url)
def prase_newslist(newslist):
    db=db_connect.init_db()
    i2=0
    nw=""
    for r in newslist:
        print(r)
        news_title=r['GGBT']
        news_ID=r['GGDM']
        news_DATE=r['FBSJ']
        news_LMID=r['LMDM']
        news_LMNAME=r['LMMC']
        news_FBBM=r['FBBM']
        if int(r['CONMENT_TYPE'])==2:
            news_url=r['GO_URL']
            news_content={"news_content":"阅读新闻：<a href='{0}'>{1}</a>".format(news_url,news_title),"fj":[]}
            print(news_url)
        else:
            news_url="https://ehall.nbu.edu.cn/new/indexnbu.html?type=5?ggdm="+news_ID
            
            #print(news_content)
            #检测文章ID是否存在
        if db_connect.find_item(db,news_ID)==0:
            b=""
            
            print("新闻不存在："+news_title+",开始抓取内容")
            if int(r['CONMENT_TYPE'])==1:
                news_content=prase_news(news_ID)
                b="<br><br>阅读新闻：<a href='{0}'>{1}</a>".format(news_url,news_title)
                news_content['news_content']=news_content['news_content']+b
            if config.fj_flag:
                if db_connect.find_item(db,news_ID)==0:
                    #下载附件
                    for rs in news_content['fj']:
                        if not os.path.exists('fj'):
                            os.mkdir('fj')
                        rs_url="https://ehall.nbu.edu.cn/"+rs['fjurl']
                        rs_name=rs['fjname']
                        print("抓取附件："+rs_name+"")
                        download_file('fj/'+rs_name,rs_url)
            else:
                for rs in news_content['fj']:
                    rs_url="https://ehall.nbu.edu.cn/"+rs['fjurl']
                    rs_name=rs['fjname']
                    news_content['news_content']=news_content['news_content']+"<br>下载附件：<a href='{0}'>{1}</a>".format(rs_url,rs_name)
            i={"news_ID":news_ID,"news_title":news_title,"news_DATE":news_DATE,"news_LMID":news_LMID,"news_LMNAME":news_LMNAME,"news_FBBM":news_FBBM,"news_url":news_url,"news_FJ":json.dumps(news_content['fj']),"news_content":news_content['news_content']}
            db_connect.add_item(db,i)
            #后续推送代码
            if config.news_pusher_type=="email" and config.email_send_type=="smtp":
                print("正在使用邮件方式推送...")
                push.push_through_email(i)
            if config.news_pusher_type=="email" and config.email_send_type=="mailgun":
                print("正在使用MailGun方式推送...")
                push.push_through_mailgun(i)
        if config.news_pusher_type=="emaillist":
            nw=nw+"{0}.<a href='{1}'>{2}</a><span style='margin-left:15px'>发布于：{3}</span><br><br>".format(str(i2+1),news_url,news_title,news_DATE)
            i2=i2+1
        else:
            print("新闻已存在："+news_title)
    if i2>0:
        nw="<p>以下是截止到{}的最新新闻列表：</p><br><br>".format(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))+nw
        if config.news_pusher_type=="emaillist":
            if config.email_send_type=="smtp":
                print("正在使用邮件列表方式推送...")
                push.push_through_email({"news_title":"新闻列表：{}".format(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())),"news_content":nw})
            if config.email_send_type=="mailgun":
                print("正在使用邮件列表方式推送...")
                push.push_through_mailgun({"news_title":"新闻列表：{}".format(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())),"news_content":nw})
                

def prase_news(newsid):
    if not config.content_flag:
        news_url="https://ehall.nbu.edu.cn/new/indexnbu.html?type=5?ggdm="+newsid
        return {"news_content":"阅读新闻：<a href='{0}>阅读新闻</a>".format(news_url),"fj":[]}
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
                print('有附件...抓取')
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



