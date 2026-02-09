# 调试与错误处理

## 调试模式

### 开启调试

```vb
Public DebugStart As Boolean
```

**示例：**
```vb
Dim http As New cHttpClient
http.DebugStart = True  ' 开启调试模式

http.SendGet("https://api.example.com/data")

' 输出调试信息
Debug.Print http.DebugInfo.Encode(, , True)
```

### 调试信息结构

`DebugInfo` 是 `cJson` 对象，包含以下字段：

```json
{
  "Request": {
    "IsAsync": false,
    "Method": "GET",
    "Url": "https://api.example.com/data",
    "Body": "",
    "Headers": {...},
    "TimeOut": 5,
    "ChartSet": "utf-8"
  },
  "Response": {
    "Status": 200,
    "StatusText": "OK",
    "Headers": {...},
    "Content": "..."
  },
  "Error": {
    "Number": 0,
    "Source": "",
    "Description": ""
  }
}
```

## 错误处理

### LastError 属性

```vb
Public LastError As String
```

存储最后一次错误的描述信息。

**示例：**
```vb
http.SendGet("https://api.example.com/data")
If http.LastError <> "" Then
    Debug.Print "错误: " & http.LastError
End If
```

### 异常信息格式

当发生 4xx/5xx 错误时，异常描述格式为：

```
状态码#状态文本#响应内容前1024字符
```

**示例：**
```vb
On Error GoTo ErrorHandler

http.SendGet("https://api.example.com/not-found")
' 抛出异常: 404#Not Found#{"error": "Resource not found"}

Exit Sub

ErrorHandler:
    ' 解析错误信息
    Dim parts() As String
    parts = Split(Err.Description, "#")
    
    If UBound(parts) >= 2 Then
        Debug.Print "HTTP 状态码: " & parts(0)
        Debug.Print "状态文本: " & parts(1)
        Debug.Print "响应内容: " & parts(2)
    Else
        Debug.Print "错误: " & Err.Description
    End If
```

## 常见错误代码

| 错误代码 | 含义 | 解决方案 |
|----------|------|----------|
| 500 | 请求的URL地址为空 | 检查 URL 参数 |
| 900 | 请求超时 | 增加 RequestTimeOut |
| 310 | 重定向次数超过限制 | 检查重定向链或增加 MaxRedirects |
| 4xx | 客户端错误 | 检查请求参数和认证 |
| 5xx | 服务器错误 | 检查服务端状态 |

## 超时错误

### 设置超时时间

```vb
http.RequestTimeOut = 60  ' 60 秒
```

### 超时处理

```vb
On Error GoTo ErrorHandler

http.RequestTimeOut = 5
http.SendGet("https://slow-api.example.com/data")

Exit Sub

ErrorHandler:
    If Err.Number = 900 Then
        Debug.Print "请求超时，请稍后重试"
    Else
        Debug.Print "其他错误: " & Err.Description
    End If
```

## 网络错误处理

```vb
Sub SafeRequest()
    Dim http As New cHttpClient
    
    On Error GoTo ErrorHandler
    
    http.DebugStart = True
    http.RequestTimeOut = 30
    
    http.SendGet("https://api.example.com/data")
    
    ' 处理响应
    ProcessResponse http.ReturnJson()
    
    Exit Sub
    
ErrorHandler:
    Dim errorMsg As String
    
    Select Case Err.Number
        Case 900
            errorMsg = "请求超时，请检查网络连接"
        Case 500
            errorMsg = "URL 地址不能为空"
        Case 404
            errorMsg = "请求的资源不存在"
        Case 401, 403
            errorMsg = "权限不足，请检查认证信息"
        Case Else
            errorMsg = "请求失败: " & Err.Description
    End Select
    
    Debug.Print errorMsg
    
    ' 输出详细调试信息
    If Not http.DebugInfo Is Nothing Then
        Debug.Print "详细调试信息:"
        Debug.Print http.DebugInfo.Encode(, , True)
    End If
End Sub
```

## 日志记录

### 记录完整请求日志

```vb
Sub LogRequest()
    Dim http As New cHttpClient
    http.DebugStart = True
    
    http.SendGet("https://api.example.com/data")
    
    ' 保存到日志文件
    Dim logFile As String
    logFile = "C:\Logs\http_" & Format(Now, "yyyymmdd_hhmmss") & ".json"
    
    Open logFile For Output As #1
    Print #1, http.DebugInfo.Encode(, , True)
    Close #1
    
    Debug.Print "日志已保存到: " & logFile
End Sub
```

## 调试技巧

### 1. 比较请求和响应

```vb
With http.DebugInfo
    Debug.Print "请求 URL: " & .Item("Request").Item("Url")
    Debug.Print "响应状态: " & .Item("Response").Item("Status")
End With
```

### 2. 检查请求头

```vb
Dim headers As Object
Set headers = http.DebugInfo.Item("Request").Item("Headers")

Dim key As Variant
For Each key In headers.Keys
    Debug.Print key & ": " & headers(key)
Next
```

### 3. 验证响应内容

```vb
Dim responseText As String
responseText = http.DebugInfo.Item("Response").Item("Content")

If Len(responseText) > 1000 Then
    Debug.Print "响应内容过长，已截断"
End If
```
