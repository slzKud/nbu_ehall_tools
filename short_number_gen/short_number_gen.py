import unifined_login,config,requests,db_connect,sqlite3,json
session=unifined_login.auto_login()
def get_data(q_str,totalsize,pagesize,pagestart):
    cur=db_connect.init_db() 
    session=unifined_login.cookies_login()
    pageNumber=int(totalsize/pagesize)
    if totalsize-pageNumber*pagesize>0:
        pageNumber=pageNumber+1
    for i in range(pagestart,pageNumber+1):
        s="querySetting={2}&pageSize={0}&pageNumber={1}".format(pagesize,i,q_str)
        print('正在获取{}页的数据...'.format(i))
        p=session.post('https://ehall.nbu.edu.cn/nbuapp/sys/lxfs/modules/jzgdhcx/V_LXFSCX_JZGDH_QUERY.do',data=s)
        if p.status_code==200:
            p1=json.loads(p.text)
            if p1['code']=="0":
                print('获取{}页的数据成功...'.format(i))
                p2=p1['datas']['V_LXFSCX_JZGDH_QUERY']['rows']
                for p3 in p2:
                    if db_connect.find_item(cur,p3['XM'],p3['DH'],p3['BM'])==0:
                        print('添加:{0}/{1}...'.format(p3['XM'],p3['BM']))
                        db_connect.add_item(cur,p3)
                    else:
                        print('跳过:{0}/{1}...'.format(p3['XM'],p3['BM']))
    print('处理完成！')
headers = {
        'Accept': 'application/json, text/plain, */*',
        'User-Agent': 'Mozilla/5.0 (Linux; Android 4.4.4; OPPO R11 Plus Build/KTU84P) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/33.0.0.0 Safari/537.36 yiban/8.1.11 cpdaily/8.1.11 wisedu/8.1.11',
        'content-type': 'application/json',
        'Accept-Encoding': 'gzip,deflate',
        'Accept-Language': 'zh-CN,en-US;q=0.8',
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'Referer': 'https://ehall.nbu.edu.cn/nbuapp/sys/lxfs/*default/index.do'
}
session.get('https://ehall.nbu.edu.cn/nbuapp/sys/lxfs/*default/index.do#/jzgdhcx')
p=session.post('https://ehall.nbu.edu.cn/nbuapp/code/f2b0992b-4e0a-471d-b678-e37a25190f28.do',headers=headers)
print(p.text)
if p.status_code==200:
        print(p.text)
        p1=json.loads(p.text)
        if p1['code']=="0":
            p2=p1['datas']['code']['rows']
            for p33 in p2:
                j=[{"name":"BM","caption":"部门","linkOpt":"AND","builderList":"cbl_m_List","builder":"m_value_equal","value":p33['id'],"value_display":p33['name']}]
                j1=json.dumps(j)
                pp=session.post('https://ehall.nbu.edu.cn/nbuapp/sys/lxfs/modules/jzgdhcx/V_LXFSCX_JZGDH_QUERY.do',data='querySetting={}&pageSize=20&pageNumber=1'.format(j1))
                if pp.status_code==200:
                    p1=json.loads(pp.text)
                    if p1['code']=="0":
                        print('获取初始数据成功！')
                        totalsize=p1['datas']['V_LXFSCX_JZGDH_QUERY']['totalSize']
                        pageNumber=1
                        pagesize=p1['datas']['V_LXFSCX_JZGDH_QUERY']['pageSize']
                        get_data(j1,totalsize,pagesize,1)
#{"POST":{"scheme":"https","host":"ehall.nbu.edu.cn","filename":"/nbuapp/code/f2b0992b-4e0a-471d-b678-e37a25190f28.do","remote":{"地址":"210.33.16.122:443"}}}
#https://ehall.nbu.edu.cn/nbuapp/code/f2b0992b-4e0a-471d-b678-e37a25190f28.do 