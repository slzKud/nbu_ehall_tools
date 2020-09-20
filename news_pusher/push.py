import getpass,json,os,config
import smtplib,requests
from email.mime.text import MIMEText
from email.utils import formataddr
def get_the_smtp_send():
    if not os.path.exists('config_smtp.json'):
        username=input("请输入发送的电子邮件账户:")
        password=getpass.getpass('请输入密码/授权码：')
        server=input("请输入SMTP服务器地址:")
        server_port=input("请输入SMTP服务器端口[25]:")
        address=input('请输入推送的电子邮件地址,以逗号分隔：').split(',')
        e={"username":username,"password":password,"address":address,"server":server,"server_port":server_port}
        es=json.dumps(e)
        fiobj=open('config_smtp.json',"w")
        fiobj.write(es)
        fiobj.close()
        return e
    else:
        if config.content_flag==False:
            return False
        fiobj=open('config_smtp.json',"r")
        es=fiobj.read()
        fiobj.close()
        e=json.loads(es)
        return e
def get_the_mailgun_send():
    if not os.path.exists('config_mailgun.json'):
        username=input("请输入发送的mailgun域名:")
        password=getpass.getpass('请输入API_KEY：')
        address=input('请输入推送的电子邮件地址,以逗号分隔：').split(',')
        e={"username":username,"password":password,"address":address}
        es=json.dumps(e)
        fiobj=open('config_mailgun.json',"w")
        fiobj.write(es)
        fiobj.close()
        return e
    else:
        if config.content_flag==False:
            return False
        fiobj=open('config_mailgun.json',"r")
        es=fiobj.read()
        fiobj.close()
        e=json.loads(es)
        return e
def push_through_email(i):
    i1=get_the_mailgun_send()
    if config.content_flag==False:
        return False
    ret=True
    try:
        print("正在发送邮件...")
        msg=MIMEText('填写邮件内容','plain','utf-8')
        msg['From']=formataddr(["今日新闻",i1['username']])  # 括号里的对应发件人邮箱昵称、发件人邮箱账号
        p=",".join(i1['address'])
        msg['To']=p            # 括号里的对应收件人邮箱昵称、收件人邮箱账号
        msg['Subject']="{}".format(i['news_title'])                # 邮件的主题，也可以说是标题
        server=smtplib.SMTP_SSL(i1['server'], int(i1['server_port']))  # 发件人邮箱中的SMTP服务器，端口是25
        server.login(i1['username'], i1['password'])  # 括号中对应的是发件人邮箱账号、邮箱密码
        server.sendmail(i1['username'],i1['address'],i['news_content'].encode('utf-8'))  # 括号中对应的是发件人邮箱账号、收件人邮箱账号、发送邮件
        server.quit()  # 关闭连接
    except:
        ret=False
    return ret
def push_through_mailgun(i):
    i1=get_the_mailgun_send()
    if config.content_flag==False:
        return False
    print("正在发送邮件...")
    f=[]
    if config.fj_flag:
        fj=json.loads(i['news_FJ'])
        for fjchild in fj:
            rs_name=fjchild['fjname']
            filename="fj/"+rs_name
            f.append(("attachment", (rs_name, open(filename,"rb").read())))
    A=requests.post(
		"https://api.mailgun.net/v3/{}/messages".format(i1['username']),
		auth=("api", i1['password']),
        files=f,
		data={"from": "News Bot <mailgun@{}>".format(i1['username']),
			"to": i1['address'],
			"subject": i['news_title'],
			"html": i['news_content'].encode('utf-8')})
    print(A.text)
    return(A.text)