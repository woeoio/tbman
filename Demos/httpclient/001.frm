VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12750
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   12750
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   615
      Left            =   7440
      TabIndex        =   4
      Top             =   6600
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   615
      Left            =   4560
      TabIndex        =   3
      Top             =   6600
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "json"
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   6600
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   9360
      TabIndex        =   1
      Top             =   6720
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   6135
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   240
      Width           =   12375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const ApiHost$ = "http://cloud-workorder-pc.test.eslink.net.cn"




Private Sub Command1_Click()
    
    With New cHttpClient
        .RequestHeaders.Add "Accept", "application/json"
        .RequestHeaders.Add "Content-Type", "application/json;charset=utf-8"
        .RequestHeaders.Add "AccessToken", "5021a00d60db149fa8931f66e2f9d854"
        .RequestHeaders.Add "Authorization", "20240928101610"
        Dim i As Long
        With New cJson
            .Item("sysStuffCode") = "TEST001"
            .Item("quantity") = 2
            With .NewItems("detailList")
                For i = 0 To 3
                    With .NewItem()
                        .Item("test") = 123
                        .Item("time") = Now()
                    End With
                Next
            End With
            'detailList 数组结束
            Dim PostBody As String: PostBody = .Encode()
        End With
        Dim Url As String: Url = ApiHost & "/workorder/out/warehouse/person/receive"
        Text1.Text = .Fetch(ReqPost, Url, PostBody).ReturnJson().Encode(, 2, True)
    End With
    
End Sub

Private Sub Command2_Click()
    MsgBox VBMAN.Version
    MsgBox VBMAN.HttpClient.Fetch(ReqGet, "http://a-vi.com/hello/VBMAN").ReturnText()
    With VBMAN.Json
        .Item("a") = 1
        .Item("b") = "dengwei"
        With .NewItems("c")
            Dim i As Long
            For i = 0 To 3
                With .NewItem()
                    .Item("d") = Now()
                    .Item("e") = 34 + i
                    .Item("f") = "进入坦克: " & i
                End With
            Next
        End With
        MsgBox .Encode(, 2, True)
        .SaveTo "c:\tmp\VBMAN_demo.json", , 2
    End With
    
    Dim Test As New VBMANLIB.cJson:  Call Test.LoadFrom("c:\tmp\VBMAN_demo.json")
    MsgBox Test("c")(2)("d"), , Test("b")
    '循环读取json数组
    Dim x As Variant
    For Each x In Test("c")
        MsgBox "for each x = " & x("f")
    Next
    '循环读取json数组
    Dim ii As Long
    For ii = 1 To Test("c").Count
        MsgBox "for i = " & Test("c")(ii)("f")
    Next
    '单独读取指定数组（下标从1开始）
    MsgBox VBMAN.Json.LoadFrom("c:\tmp\VBMAN_demo.json")("c")(2)("f")
    
End Sub

Private Sub Command3_Click()
    Text1.Text = VBMAN.Json.LoadFrom("c:\tmp\VBMAN_demo.json").Encode(, 2)
    MsgBox "json 文件已加载，并转为字符串显示在文本框"
    Dim C As New cJson
    '文本框内容解析为 json 对象
    '（本质就是字典对象，可以直接用，其中数组是集合对象， 所以下标从1开始）
    C.Decode Text1.Text
    '现在可以使用了
    MsgBox "b = " & C("b")
    MsgBox "第一个数组的 f 值 = " & C("c")(1)("f")
    '现在修改 a 节点
    C("a") = "邓伟”"
    '把修改后的 json 对象编码为字符串并显示到文本框,
    '其中第二个参数表示用 n 个空格来格式化字符串，如果不写，字符串会缩成一团
    '第三个参数为 true 则会显示中文, 而不是 uncode 码
    Text1.Text = C.Encode(, 2, True)
End Sub

Private Sub Command4_Click()
    '从文件加载 json 内容并使用
    With VBMAN.Json.LoadFrom("c:\tmp\VBMAN_demo.json")
        MsgBox .Root("b")
        MsgBox .Root("c")(1)("f")
    End With
    '从网络加载 json 内容并使用
    With VBMAN.HttpClient.Fetch(ReqGet, "http://a-vi.com/hello/VBMAN").ReturnJson()
        MsgBox .Root("name")                                                    '显示name字段
        MsgBox .Root("time")                                                    '显示time字段
        MsgBox .Encode()                                                        '显示收到的完整内容
    End With
End Sub
