# 响应处理

## 获取响应文本

### ReturnText 方法

```vb
Public Function ReturnText(Optional IsUtf8 As Boolean = True, Optional IsConvert As Boolean) As String
```

| 参数 | 说明 | 默认值 |
|------|------|--------|
| IsUtf8 | 使用 UTF-8 解码 | True |
| IsConvert | 使用 StrConv 转换（解决乱码） | False |

**示例：**
```vb
' 默认 UTF-8 解码
Dim text As String
text = http.ReturnText()

' 使用系统编码
text = http.ReturnText(False)

' 强制转换（处理乱码）
text = http.ReturnText(False, True)
```

## 获取 JSON 数据

### ReturnJson 方法

```vb
Public Function ReturnJson(Optional IsUtf8 As Boolean = True, Optional IsConvert As Boolean) As cJson
```

**示例：**
```vb
http.SendGet("https://api.example.com/users/1")

Dim json As cJson
Set json = http.ReturnJson()

Debug.Print json.Item("name")
Debug.Print json.Item("email")
```

## 获取二进制数据

### ReturnBody 方法

```vb
Public Function ReturnBody() As Byte()
```

**示例：**
```vb
Dim bytes() As Byte
bytes = http.ReturnBody()

' 保存到文件
Open "C:\data.bin" For Binary As #1
Put #1, , bytes
Close #1
```

## 获取流数据

### ReturnStream 方法

```vb
Public Function ReturnStream() As Variant
```

**示例：**
```vb
Dim stream As Variant
Set stream = http.ReturnStream()
```

## 响应头信息

### ResponseHeaders 字典

```vb
Public ResponseHeaders As New Scripting.Dictionary
```

**示例：**
```vb
http.SendGet("https://api.example.com/data")

' 获取所有响应头
Dim key As Variant
For Each key In http.ResponseHeaders.Keys
    Debug.Print key & ": " & http.ResponseHeaders(key)
Next

' 获取特定响应头
Dim contentType As String
If http.ResponseHeaders.Exists("Content-Type") Then
    contentType = http.ResponseHeaders("Content-Type")
End If
```

### 获取 Set-Cookie

```vb
' WinHttp 会自动处理 Cookie，但可以通过响应头获取
If http.ResponseHeaders.Exists("Set-Cookie") Then
    Debug.Print http.ResponseHeaders("Set-Cookie")
End If
```

## 状态码处理

### 获取状态信息

```vb
http.SendGet("https://api.example.com/data")

Debug.Print "状态码: " & http.Inst.Status
Debug.Print "状态文本: " & http.Inst.StatusText
```

### 状态码说明

| 范围 | 含义 | 处理方式 |
|------|------|----------|
| 200-299 | 成功 | 正常返回 |
| 300-399 | 重定向 | 自动跟随或手动处理 |
| 400-499 | 客户端错误 | 抛出异常 |
| 500-599 | 服务器错误 | 抛出异常 |

### 错误响应处理

```vb
On Error GoTo ErrorHandler

http.SendGet("https://api.example.com/not-found")

' 如果是 4xx/5xx 会抛出异常
Debug.Print http.ReturnText()

Exit Sub

ErrorHandler:
    ' Err.Number 包含状态码
    ' Err.Description 包含 "状态码#状态文本#响应内容片段"
    Dim parts() As String
    parts = Split(Err.Description, "#")
    
    Debug.Print "状态码: " & parts(0)
    Debug.Print "状态文本: " & parts(1)
    Debug.Print "响应内容: " & parts(2)
```

## Cookies 管理

### Cookies 字典

```vb
Public Cookies As New Scripting.Dictionary
```

**示例：**
```vb
' 设置 Cookie
http.SetCookies("session=abc123; path=/")

' 读取 Cookie（需自行解析）
Dim cookieValue As String
If http.Cookies.Exists("session") Then
    cookieValue = http.Cookies("session")
End If
```

## 响应数据缓存

### ResponseRaw 属性

```vb
Public ResponseRaw As Variant
```

用于缓存原始响应内容。

## 完整响应处理示例

```vb
Sub HandleResponse()
    Dim http As New cHttpClient
    
    http.SendGet("https://api.example.com/users/1")
    
    ' 1. 检查状态码
    Select Case http.Inst.Status
        Case 200
            Debug.Print "请求成功"
        Case 201
            Debug.Print "创建成功"
        Case 204
            Debug.Print "无内容返回"
        Case Else
            Debug.Print "其他状态: " & http.Inst.Status
    End Select
    
    ' 2. 处理 Content-Type
    Dim contentType As String
    If http.ResponseHeaders.Exists("Content-Type") Then
        contentType = http.ResponseHeaders("Content-Type")
        
        If InStr(contentType, "application/json") > 0 Then
            ' JSON 响应
            Dim json As cJson
            Set json = http.ReturnJson()
            ProcessJsonResponse json
        ElseIf InStr(contentType, "text/") > 0 Then
            ' 文本响应
            ProcessTextResponse http.ReturnText()
        Else
            ' 二进制响应
            ProcessBinaryResponse http.ReturnBody()
        End If
    End If
    
    ' 3. 获取 Cookie
    If http.ResponseHeaders.Exists("Set-Cookie") Then
        SaveCookies http.ResponseHeaders("Set-Cookie")
    End If
End Sub
```
