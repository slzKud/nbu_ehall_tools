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
#email以电子邮件发送（每个新闻一个右键）（暂不可用）
#emaillist以电子邮件列表发送（一次新闻列表右键）（暂不可用）
#tgbot以电报机器人信息发送（暂不可用）
news_pusher_type="command" 
#电子邮件设置（挖坑）
#电报机器人设置（挖坑）