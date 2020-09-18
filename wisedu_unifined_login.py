import requests,config,getpass,json
username=input("学号:")
password=getpass.getpass('密码：')
url=config.unifiedloginapi
url=url.replace('$username$',username)
url=url.replace('$password$',password)
session=requests.session()
p=session.get(url)
if p.status_code==200:
    p1=json.loads(p.text)
    if p1['code']==0:
        print('登陆成功！')
        c=p1['cookies']
        fiobj=open('uis_nbu_cookies.txt',"w")
        fiobj.write(c)
        fiobj.close
        print('写入成功！')
    else:
        print('登陆失败！：'+p1['msg'])