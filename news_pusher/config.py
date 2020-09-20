#wisedu_unifined_login
#https://github.com/ZimoLoveShuang/wisedu-unified-login-api
#登陆服务器地址
unifiedloginapi="http://www.zimo.wiki:8080/wisedu-unified-login-api-v1.0/api/login?login_url=https%3A%2F%2Fuis.nbu.edu.cn%2Fauthserver%2Flogin&password=$password$&username=$username$"
#新闻抓取深度
page_deepth=3 #只抓取前三页内容
page_size=15 #每页15个
#抓取设置
fj_flag=False #抓取附件设置 False代表不抓取
content_flag=True #抓取新闻内容 False代表不抓取
#储存设置
db_file='news.db' #储存的数据库名
#推送设置
#command以命令行呈现
#email以电子邮件发送（每个新闻一个邮件）
#emaillist以电子邮件列表发送（一次新闻列表邮件）（暂不可用）
#tgbot以电报机器人信息发送（暂不可用）
#sms以腾讯云手机信息发送（暂不可用）
news_pusher_type="email" 
#电子邮件设置（挖坑）
#邮件发送类型
#smtp:使用smtp发送邮件，大规模发送可能会被认为是SPAM而发送失败
#mailgun:使用mailgun发送邮件，大规模发送收费
email_send_type="smtp"
#电报机器人设置（挖坑）
#手机信息设置（挖坑）