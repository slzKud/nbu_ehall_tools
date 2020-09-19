# news_pusher
这是一个用于抓取并储存宁大新闻网的新闻的工具，并可以选择方式进行推送
**本工具尚未完成**
## 如何使用
1. 安装依赖`request`，并修改`config.py`
2. 运行脚本`news_pusher.py`,第一次使用会输入学号和密码用于登陆
3. 抓取完成后会把新闻储存在`news.db`,附件存放在`fj`文件夹
## 配置文件默认配置
```
#新闻抓取深度

page_deepth=3 #只抓取前三页内容

page_size=15 #每页15个

#抓取设置

fj_flag=False #抓取附件设置 False代表不抓取

content_flag=True #抓取新闻内容 False代表不抓取

#储存设置

db_file='news.db' #储存的数据库名</code>
```