VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "请在此登陆"
   ClientHeight    =   6330
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   4425
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command1 
      Caption         =   "放弃获取"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Menu T1 
      Caption         =   "js回调测试"
      Visible         =   0   'False
   End
   Begin VB.Menu T2 
      Caption         =   "Cookies获取"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mb_callback As MiniblinkCallBack
Attribute mb_callback.VB_VarHelpID = -1
Private mb_api As New MiniblinkAPI
Private Declare Function SetProcessDpiAwarenessContext Lib "user32" (dpi As Long) As Boolean
Private mb As Long

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Activate()
    Me.ScaleMode = 3
    
    mb_api.wkeInitializeEx 0
    
    mb_api.wkeJsBindFunction "test", mb_callback.wkeJsNativeFunction, 0, 2               'js回调事件绑定（影响所有webview和webwindow）
    
    mb = mb_api.wkeCreateWebWindow(2, Me.hWnd, 0, 0, Me.ScaleWidth, Me.ScaleHeight)
    mb_api.wkeShowWindow mb, True
    
    mb_api.wkeOnLoadUrlBegin mb, mb_callback.wkeLoadUrlBeginCallback, 0                  'url加载事件绑定
    mb_api.wkeOnCreateView mb, mb_callback.wkeCreateViewCallback, 0                      '创建新窗口事件绑定
    mb_api.wkeOnDownload mb, mb_callback.wkeDownloadCallback, 0 '下载事件绑定
    mb_api.wkeSetUserAgent mb, "Mozilla/5.0 (iPhone; CPU iPhone OS 10_3 like Mac OS X) AppleWebKit/602.1.50 (KHTML, like Gecko) CriOS/56.0.2924.75 Mobile/14E5239e Safari/602.1"
    'mb_api.wkeSetCookieJarFullPath mb, App.Path & "\cookies.dat"
    mb_api.wkeLoadURL mb, "https://uis.nbu.edu.cn/authserver/login?service=http%3A%2F%2Fehall.nbu.edu.cn%2Flogin%3Fservice%3Dhttps%3A%2F%2Fehall.nbu.edu.cn%2Fnew%2Findex.html%3Fbrowser%3Dno"
    'mb_api.wkeLoadHTML mb, "<h1>test</h1>"
End Sub

Private Sub Form_Load()
    If ifWin101607 Then SetProcessDpiAwarenessContext 3&
    Set mb_callback = New MiniblinkCallBack
    'Kill App.Path & "\cookies.dat"
End Sub


Private Sub mb_callback_wkeDocumentReadyCallback(ByVal webView As Long, ByVal param As Long)
Me.Show
End Sub

Private Sub mb_callback_wkeDownloadCallback(ByVal webView As Long, ByVal param As Long, ByVal url As String)
    Debug.Print "触发了下载事件，下载地址：" & url
End Sub

Private Sub mb_callback_wkeJsNativeFunction(ByVal es As Long, ByVal param As Long)
    Dim tret1 As Currency, tret2 As Currency
    tret1 = mb_api.jsArg(es, 0)
    tret2 = mb_api.jsArg(es, 1)
    MsgBox mb_api.jsToTempStringW(es, tret1) & "/" & mb_api.jsToTempStringW(es, tret2)
End Sub

Private Sub mb_callback_wkeLoadUrlBeginCallback(ByVal webView As Long, ByVal param As Long, ByVal url As String, ByVal job As Long)
    Debug.Print url
    If InStr(url, "https://uis.nbu.edu.cn/authserver") = 0 Then T2_Click: End
    'If InStr(url, "http://ehall.nbu.edu.cn/login?service=https://ehall.nbu.edu.cn/new/index.html") > 0 Then T2_Click: End
End Sub

Private Sub mb_callback_wkeNavigationCallback(ByVal webView As Long, ByVal param As Long, ByVal navigationType As MiniblinkSDK_200.wkeNavigationType, ByVal url As String)
 Debug.Print "触发了wkeCreateViewCallback"
    mb_callback.Return_wkeCreateViewCallback = webView      '使用原webview加载
End Sub

Private Sub T1_Click()
    mb_api.wkeRunJSW mb, "alert('test');window.test('xcv','hj自行车5gj');"
End Sub

Private Sub T2_Click()
Dim c As String
c = mb_api.wkeGetCookieW(mb)
    If c <> "" And c <> UEFLoadTextFile(App.Path & "\uis_nbu_cookies.txt", uef_utf8NB) Then
        UEFSaveTextFile App.Path & "\uis_nbu_cookies.txt", mb_api.wkeGetCookieW(mb), False, uef_utf8NB, uef_utf8NB
    End If
End Sub

Private Function ifWin101607() As Boolean
Dim a As String, B As Long
On Error Resume Next
a = get_path_from_reg("", "winbuild")
B = 0
B = CLng(a)
If B >= 14393 Then ifWin101607 = True
End Function

