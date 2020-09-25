# nbu_ehall_tools
一个用于个人用途的宁波大学电子办事大厅的辅助工具

基于python制作
## 初步完成的子模块
* uis_login

    这是一个基于miniblink的工具，用于登陆获取电子办事大厅的cookies
* class_ics_gen 

    将学期课程表导出为ics日程文件，供手机自带日历/Outlook使用

    **支持本部及梅山**
* news_pusher

    推送学校办事大厅的新闻，支持以电子邮件等形式自动推送或储存在数据库内
    
    **现只支持电子邮件推送（含内容和附件或发送邮件列表）**
* short_number_get

    教职工短号抓取和查询

    **查询暂未支持**
## 计划完成的部分

* exam_ics_gen

    将考试安排导出为ics日程文件，供手机自带日历/Outlook使用

* exam_score_pusher

    定时查询成绩，并推送你的成绩
* cal_viewer

    免登录查看校历，并通过OCR内容获取关键信息


## 如何使用
* class_ics_gen 
    1. 安装依赖ics和requests

        `pip install ics requests`

    2. 使用[wisedu-unified-login-api](https://github.com/ZimoLoveShuang/wisedu-unified-login-api)或[uis_login](https://github.com/slzKud/nbu_ehall_tools/tree/master/uis_login)登陆系统，并保存cookies到`uis_nbu_cookies.txt`
    3. 运行`class_ics_gen.py`,稍等片刻后生成课程表ics文件。可在Windows系统自带的日历上使用
* news_pusher

    请参考子项目文件夹的Readme
* uis_login

    请参考子项目文件夹的Readme

