import requests,config,getpass,json,os
def go_login(username,password):
    url=config.unifiedloginapi
    url=url.replace('$username$',username)
    url=url.replace('$password$',password)
    session=requests.session()
    p=session.get(url)
    if p.status_code==200:
        p1=json.loads(p.text)
        if p1['code']==0:
            c=p1['cookies']
            print('登陆成功！')
            fiobj=open('uis_nbu_cookies.txt',"w")
            fiobj.write(c)
            fiobj.close
            cookieStr=c
            cookies = {}
            for line in cookieStr.split(';'):
                name, value = line.strip().split('=', 1)
                cookies[name] = value
            session.cookies = requests.utils.cookiejar_from_dict(cookies)
            session.headers= {
                'Accept': 'application/json, text/plain, */*',
                'User-Agent': 'Mozilla/5.0 (Linux; Android 4.4.4; OPPO R11 Plus Build/KTU84P) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/33.0.0.0 Safari/537.36 yiban/8.1.11 cpdaily/8.1.11 wisedu/8.1.11',
                'content-type': 'application/json',
                'Accept-Encoding': 'gzip,deflate',
                'Accept-Language': 'zh-CN,en-US;q=0.8',
                'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8'
            }
            return session
        else:
            print('登陆失败！：'+p1['msg'])
            return False
def auto_login():
    if not os.path.exists('config_user.json'):
        username=input("学号:")
        password=getpass.getpass('密码：')
    else:
        fiobj=open('config_user.json',"r")
        j=fiobj.read()
        fiobj.close()
        jobj=json.loads(j)
        username=jobj['username']
        password=jobj['password']
    a=go_login(username,password)
    if a!=False:
        if not os.path.exists('config_user.json'):
            jobj={'username':username,"password":password}
            j=json.dumps(jobj)
            fiobj=open('config_user.json',"w")
            fiobj.write(j)
            fiobj.close()
        return a
def cookies_login():
    c=""
    if not os.path.exists('uis_nbu_cookies.txt'):
        print("cookies为空，即将打开cookies获取程序")
        return False
    fiobj=open('uis_nbu_cookies.txt',"r")
    c=fiobj.read()
    fiobj.close()
    if c=="":
        print("cookies为空")
        return False
    print("cookies加载成功!")
    cookieStr=c
    cookies = {}
    for line in cookieStr.split(';'):
        name, value = line.strip().split('=', 1)
        cookies[name] = value
    session = requests.session()
    session.cookies = requests.utils.cookiejar_from_dict(cookies)
    session.headers = {
            'Accept': 'application/json, text/plain, */*',
            'User-Agent': 'Mozilla/5.0 (Linux; Android 4.4.4; OPPO R11 Plus Build/KTU84P) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/33.0.0.0 Safari/537.36 yiban/8.1.11 cpdaily/8.1.11 wisedu/8.1.11',
            'content-type': 'application/json',
            'Accept-Encoding': 'gzip,deflate',
            'Accept-Language': 'zh-CN,en-US;q=0.8',
            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8'
    }
    return(session)
    
