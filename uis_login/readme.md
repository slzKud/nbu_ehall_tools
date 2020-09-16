# uis_login
这是一个基于miniblink的工具，用于登陆获取电子办事大厅的cookies

基于[miniblink](https://github.com/weolar/miniblink49)开源版本制作，使用前请[下载](https://github.com/weolar/miniblink49/releases)核心模块和[VB6 SDK](https://github.com/imxcstar/vb6-miniblink-SDK)

注册VB6 SDK(`MiniblinkSDK_200.dll`)并复制核心DLL(`node.dll`)后使用VB6编译后运行

运行后自动打开统一身份认证登陆页面，登陆成功后自动生成`uis_nbu_cookies.txt` 内存有可以登陆电子办事大厅的cookies

## 注意
**本程序仅为实验用途，仅供个人使用**

不支持**自动登录**，如需自动化登陆可以移步[wisedu-unified-login-api](https://github.com/ZimoLoveShuang/wisedu-unified-login-api)
