VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14505
   LinkTopic       =   "Form1"
   ScaleHeight     =   9420
   ScaleWidth      =   14505
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "请求"
      Height          =   615
      Left            =   10320
      TabIndex        =   2
      Top             =   240
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      Height          =   7935
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   1080
      Width           =   14055
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Text            =   "46349936"
      Top             =   240
      Width           =   10095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const API_BASE As String = "https://bf-api.cfhy.com"
'接口列表
Const API_PULL_WAYBILL As String = API_BASE & "/api/common/pullWaybill"

Const CLIENT_ID As String = "55S1272A"

Private Sub Command1_Click()
    '构造请求数据
    Dim Body As String
    With New cJson
        .Item("markedValues") = "WGGBXT00001"
        .Item("clientId") = CLIENT_ID
        .Item("wayBillId") = Text1.Text
        .Item("token") = "cbe84888-9f48-4c00-aae6-3170bf5951cd"
        .Item("source") = "1001"
        .Item("validation") = "Y2JlODQ4ODgtOWY0OC00YzAwLWFhZTYtMzE3MGJmNTk1MWNkNTVTMTI3MkE0NjM0OTkzNg=="
        Body = .Encode()
        Debug.Print Body
    End With
    '开启错误捕获，以便输出错误调试信息
    On Error GoTo EH
    '如果需要隔离多个请求，可以new新建独立 http 客户端实例
    '    With New cHttpClient
    '这里使用全局共用的实例
    With VBMAN.HttpClient
        '开启内部调试
        .DebugStart = True
        '设置请求的 Content-Type 为 JSON
        '        .RequestHeaders.Add "Content-Type", "application/json"
        '内置了一些常见类型
        .SetRequestContentType JsonString
        '请求  post 接口 （新版vbman提供了多个send方法，比fetch少一个参数）
        .SendPost API_PULL_WAYBILL, Body
        '处理返回结果
        With .ReturnJson()
            '显示完整返回内容到文本框，后面2个参数是格式化json文本的
            Text2.Text = .Encode(, 2, True)
            '提取返回内容使用（.root 是 cjson 实例的默认成员，json的根节点对象）
            If .Root("success") = True Then
                '如果请求成功，就拿出返回的 id 和name 来使用
                MsgBox .Root("data")("cargoName"), , .Root("data")("wayBillId")
            Else
                '失败就显示错误信息提示
                MsgBox .Root("message"), , "请求失败"
            End If
        End With
    End With
    '    Exit Sub
EH:
    Debug.Print VBMAN.HttpClient.DebugInfo.Encode(, 2, True)
End Sub
