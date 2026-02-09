# 快速入门

## 基础使用

### 1. 创建实例

```vb
Dim http As New cHttpClient
```

### 2. 发送 GET 请求

```vb
' 最简单的 GET 请求
http.SendGet("https://jsonplaceholder.typicode.com/posts/1")
Debug.Print http.ReturnText()
```

### 3. 发送 POST 请求

```vb
' 表单数据
http.RequestDataForm.Add "username", "admin"
http.RequestDataForm.Add "password", "123456"
http.SendPost("https://api.example.com/login")
```

### 4. 处理 JSON 数据

```vb
' 发送 JSON
http.SetRequestContentType(Json)
http.RequestDataJson.Add "title", "Hello"
http.RequestDataJson.Add "body", "World"
http.SendPost("https://api.example.com/posts")

' 解析 JSON 响应
Dim json As cJson
Set json = http.ReturnJson()
Debug.Print json.Item("id")
```

## 链式调用

所有配置方法都返回 `cHttpClient` 实例，支持链式调用：

```vb
http.SetCookies("session=abc123") _
    .SetRequestContentType(Json) _
    .Async(False) _
    .SendGet("https://api.example.com/data")
```

## 错误处理

```vb
On Error GoTo ErrorHandler

http.SendGet("https://api.example.com/data")
Debug.Print http.ReturnText()

Exit Sub

ErrorHandler:
    Debug.Print "错误: " & Err.Description
    ' 获取详细调试信息
    If http.DebugStart Then
        Debug.Print http.DebugInfo.Encode()
    End If
```

## 完整示例

```vb
Sub TestHttpClient()
    Dim http As New cHttpClient
    
    ' 启用调试模式
    http.DebugStart = True
    
    ' 设置请求头
    http.RequestHeaders.Add "X-API-Key", "secret123"
    http.SetCookies("session=abc")
    
    ' 设置超时（秒）
    http.RequestTimeOut = 30
    
    ' 发送请求
    http.SendGet("https://httpbin.org/get")
    
    ' 输出结果
    Debug.Print "状态码: " & http.Inst.Status
    Debug.Print "响应内容: " & http.ReturnText()
    
    ' 输出调试信息
    Debug.Print http.DebugInfo.Encode(, , True)
End Sub
```
