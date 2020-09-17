import requests,ics,os,json,time,config,datetime
def local2utc(local_str):
    local_st = datetime.datetime.strptime(local_str, '%Y-%m-%d %H:%M:%S')
    time_struct = time.mktime(local_st.timetuple())
    utc_st = datetime.datetime.utcfromtimestamp(time_struct)
    utc_str=utc_st.strftime('%Y-%m-%d %H:%M:%S')
    return utc_str
c=""
if not os.path.exists('uis_nbu_cookies.txt'):
    print("cookies为空，即将打开cookies获取程序")
    os.system('uis_login.exe')
fiobj=open('uis_nbu_cookies.txt',"r")
c=fiobj.read()
fiobj.close()
if c=="":
    print("cookies为空，即将打开cookies获取程序")
    os.system('uis_login.exe')
print("cookies加载成功!")

cookieStr=c
cookies = {}
for line in cookieStr.split(';'):
        name, value = line.strip().split('=', 1)
        cookies[name] = value
session = requests.session()
session.cookies = requests.utils.cookiejar_from_dict(cookies)
headers = {
        'Accept': 'application/json, text/plain, */*',
        'User-Agent': 'Mozilla/5.0 (Linux; Android 4.4.4; OPPO R11 Plus Build/KTU84P) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/33.0.0.0 Safari/537.36 yiban/8.1.11 cpdaily/8.1.11 wisedu/8.1.11',
        'content-type': 'application/json',
        'Accept-Encoding': 'gzip,deflate',
        'Accept-Language': 'zh-CN,en-US;q=0.8',
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8'
}
#用于刷新cookies
print("正在刷新当前学期信息....")
p=session.get("https://ehall.nbu.edu.cn/jwapp/sys/wdkb/*default/index.do?t_s=1600355558352&amp_sec_version_=1&gid_=TnJwRklMUThXMkR5UW80aFJBZXhjZWhMZGhlN09HR0FaRnpsZTdyNGZuS1NPTGQxeW9oaGFjWkdwbytISzRybUovWHJ4WVhnRDJ1UVVHQkd2RmNVd1E9PQ&EMAP_LANG=zh&THEME=indigo#/xskcb",json="",headers=headers,allow_redirects=True)
print("正在获取当前学期信息....")
p=session.get("https://ehall.nbu.edu.cn/jwapp/sys/wdkb/modules/jshkcb/dqxnxq.do",json="",headers=headers,allow_redirects=True)
xqjson=p.text
xq=json.loads(xqjson)

if xq['code']=='0':
    print('获取学期信息成功!')
else:
    print('学期信息获取失败')
    exit()
if xq['datas']['dqxnxq']['totalSize']==0:
    print('学期信息获取失败:'+xq['datas']['dqxnxq']['extParams']['msg'])
    exit()
xqn=xq['datas']['dqxnxq']['rows'][0]['DM']
xna=xqn.split('-')
xndm=xna[0]+"-"+xna[1]
xqdm=xna[2]
rq=time.strftime('%Y-%m-%d')
print('当前学期：'+xq['datas']['dqxnxq']['rows'][0]['MC'])
p=session.post("https://ehall.nbu.edu.cn/jwapp/sys/wdkb/modules/xskcb/xsdkkc.do",data="XNXQDM="+xqn+"&SKZC=1&*order=-SQSJ",headers=headers,allow_redirects=True)
p=session.post("https://ehall.nbu.edu.cn/jwapp/sys/wdkb/modules/jshkcb/dqzc.do",data="XN="+xndm+"&XQ="+xqdm+"&RQ="+rq,headers=headers,allow_redirects=True)
p=session.post("https://ehall.nbu.edu.cn/jwapp/sys/wdkb/modules/jshkcb/cxjcs.do",data="XN="+xndm+"&XQ="+xqdm,headers=headers,allow_redirects=True)
p=session.post("https://ehall.nbu.edu.cn/jwapp/sys/wdkb/modules/xskcb/xswpkc.do",data="XNXQDM="+xqn+"&SKZC=1",headers=headers,allow_redirects=True)
p=session.post("https://ehall.nbu.edu.cn/jwapp/sys/wdkb/modules/xskcb/xskcb.do",data="XNXQDM="+xqn+"&SKZC=1",headers=headers,allow_redirects=True)
kcbjson=p.text
kcb=json.loads(kcbjson)

if kcb['code']=='0':
    print('获取课程信息成功!')
else:
    print('课程信息获取失败')
    exit()
if kcb['datas']['xskcb']['totalSize']==0:
    print('课程信息获取失败:'+kcb['datas']['xskcb']['extParams']['msg'])
    exit()
jtkcb=kcb['datas']['xskcb']['rows']
print('当前课程节数:'+ str(kcb['datas']['xskcb']['totalSize']))
c=ics.Calendar()
for jtkcinfo in jtkcb:
    #输出课程开始
    print("课号："+jtkcinfo['KCH']+"["+jtkcinfo['KXH']+"]")
    print("课程名："+jtkcinfo['KCM'])
    print("课程教室："+jtkcinfo['JASMC'])
    print("老师名："+jtkcinfo['SKJS'])
    print("上课日："+config.XQJ_CN[int(jtkcinfo['SKXQ'])-1])
    print("上课开始时间："+config.SKSJ_BB[int(jtkcinfo['KSJC'])-1])
    print("上课结束时间："+config.XKSJ_BB[int(jtkcinfo['JSJC'])-1])
    print("上课周次："+jtkcinfo['SKZC'])
    #处理日期
    t_str = config.XQDYZ[xqn]
    d = datetime.datetime.strptime(t_str, '%Y-%m-%d')
    d1=int(jtkcinfo['SKXQ'])-1
    skd=[]
    for j in jtkcinfo['SKZC']:
        if j=='1':
            delta = datetime.timedelta(days=d1)
            n_days = d + delta
            n_day_str=n_days.strftime('%Y-%m-%d')
            skd.append(n_day_str)
        d1=d1+7
    print("上课日：")
    print(skd)
    #添加日期
    for skt in skd:
        e=ics.Event()
        e.name=jtkcinfo['KCM']+" "+jtkcinfo['JASMC']
        e.description=jtkcinfo['KCH']+"["+jtkcinfo['KXH']+"] "+jtkcinfo['KCM']+" "+jtkcinfo['JASMC']+" "+jtkcinfo['SKJS']
        e.location=jtkcinfo['JASMC']
        #修复ics的一个bug
        b1=skt+" "+config.SKSJ_BB[int(jtkcinfo['KSJC'])-1]+":00"
        e1=skt+" "+config.XKSJ_BB[int(jtkcinfo['JSJC'])-1]+":00"
        e.begin=local2utc(b1)
        e.end=local2utc(e1)
        c.events.add(e)
fiobj=open(xqn+"课程.ics","w",encoding="utf-8-sig")
fiobj.write(str(c))
fiobj.close()
print('结果已经保存在：'+xqn+"课程.ics,处理完成")
    #输出课程结束
