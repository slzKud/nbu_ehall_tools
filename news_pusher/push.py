import getpass,json,os,config,re
import smtplib,requests
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
from email.utils import formataddr
from email.mime.multipart import MIMEMultipart
from email.header import Header
def download_file(filename,url):
    print('正在下载:'+url)
    r = requests.get(url)
    with open(filename, 'wb') as f:
        f.write(r.content)
        f.close()
    print('下载完成:'+url)
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
def dopic_smtp(content):
    msg=MIMEMultipart('alternative')
    pic_url = re.findall('src="(.*?)"',content,re.S)
    A=str(content)
    for key in pic_url:
        d=re.compile(r"http://ehall\.nbu\.edu\.cn/publicapp/sys/emapcomponent/file/getAttachmentFile/(.*)\.do")
        pic_cid=d.sub(r'\1',key)
        pic_name=pic_cid+".png"
        print(key)
        print(pic_name)
        if not os.path.exists('fj'):
            os.mkdir('fj')
        download_file('fj/'+pic_name,key)
        fp = open('fj/'+pic_name, 'rb')
        msgImage = MIMEImage(fp.read())
        fp.close()
        msgImage.add_header('Content-ID', '{}'.format(pic_cid))
        msg.attach(msgImage)
        A=A.replace(key,"cid:{}".format(pic_cid))
        print(A)
    msg.attach(MIMEText(A, 'html', 'utf-8'))
    return msg

def dopic_mailgun(content):
    pic_url = re.findall('src="(.*?)"',content,re.S)
    A=str(content)
    f=[]
    for key in pic_url:
        d=re.compile(r"http://ehall\.nbu\.edu\.cn/publicapp/sys/emapcomponent/file/getAttachmentFile/(.*)\.do")
        pic_cid=d.sub(r'\1',key)
        pic_name=pic_cid+".png"
        print(key)
        print(pic_name)
        if not os.path.exists('fj'):
            os.mkdir('fj')
        download_file('fj/'+pic_name,key)
        f.append(("inline", (pic_name, open('fj/'+pic_name,"rb").read())))
        A=A.replace(key,"cid:{}".format(pic_name))
        print(A)
    return [A,f]

def push_through_email(i):
    i1=get_the_smtp_send()
    if config.content_flag==False:
        return False
    ret=True
    print("正在发送邮件...")
    msg=dopic_smtp(i['news_content'])
    msg['From']=formataddr(["News Bot",i1['username']])  # 括号里的对应发件人邮箱昵称、发件人邮箱账号
    p=",".join(i1['address'])
    msg['To']=p            # 括号里的对应收件人邮箱昵称、收件人邮箱账号
    msg['Subject']="{}".format(i['news_title'])  
    if config.fj_flag:
        fj=json.loads(i['news_FJ'])
        for fjchild in fj:
            rs_name=fjchild['fjname']
            filename="fj/"+rs_name
            att2 = MIMEText(open(filename, 'rb').read(), 'base64', 'utf-8')
            att2["Content-Type"] = 'application/octet-stream'
            att2.add_header("Content-Disposition", "attachment", filename=("gbk", "", "{}".format(rs_name)))
            msg.attach(att2)
    server=smtplib.SMTP_SSL(i1['server'], int(i1['server_port']))  # 发件人邮箱中的SMTP服务器，端口是25
    server.login(i1['username'], i1['password'])  # 括号中对应的是发件人邮箱账号、邮箱密码
    server.sendmail(i1['username'],i1['address'],msg.as_string())  # 括号中对应的是发件人邮箱账号、收件人邮箱账号、发送邮件
    server.quit()  # 关闭连接
    return ret
def push_through_mailgun(i):
    i1=get_the_mailgun_send()
    if config.content_flag==False:
        return False
    print("正在发送邮件...")
    f1=dopic_mailgun(i['news_content'])
    f=f1[1]
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
			"html": f1[0].encode('utf-8')})
    print(A.text)
    return(A.text)