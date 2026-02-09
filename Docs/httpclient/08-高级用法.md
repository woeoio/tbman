# 高级用法

## 批量请求

### 顺序执行多个请求

```vb
Sub BatchRequests()
    Dim http As New cHttpClient
    Dim urls As Variant
    urls = Array( _
        "https://api.example.com/users/1", _
        "https://api.example.com/users/2", _
        "https://api.example.com/users/3" _
    )
    
    Dim i As Long
    For i = LBound(urls) To UBound(urls)
        http.SendGet(urls(i))
        Debug.Print "请求 " & i + 1 & ": " & http.Inst.Status
        
        If http.Inst.Status = 200 Then
            ProcessUser http.ReturnJson()
        End If
    Next
End Sub
```

### 并发请求（使用集合）

```vb
' 使用多个实例实现并发
Sub ConcurrentRequests()
    Dim http1 As New cHttpClient
    Dim http2 As New cHttpClient
    Dim http3 As New cHttpClient
    
    ' 异步发送
    http1.Async(True).SendGet("https://api.example.com/a")
    http2.Async(True).SendGet("https://api.example.com/b")
    http3.Async(True).SendGet("https://api.example.com/c")
    
    ' 等待所有请求完成...
End Sub
```

## 文件上传

### 简单文件上传（推荐）

使用内置的 `UploadFile` 方法简化文件上传：

```vb
Sub SimpleUpload()
    Dim http As New cHttpClient
    
    ' 一行代码上传文件
    If http.UploadFileSimple("https://api.example.com/upload", _
                              "C:\Documents\report.pdf") Then
        Debug.Print "上传成功!"
        Debug.Print "服务器响应: " & http.ReturnText()
    Else
        Debug.Print "上传失败: " & http.LastError
    End If
End Sub
```

### 高级文件上传

支持自定义字段名和额外表单数据：

```vb
Sub AdvancedUpload()
    Dim http As New cHttpClient
    Dim formData As New Scripting.Dictionary
    
    ' 添加额外的表单数据
    formData.Add "userId", "12345"
    formData.Add "category", "documents"
    
    ' 上传文件
    http.UploadFile "https://api.example.com/upload", _
                    "C:\Documents\report.pdf", _
                    fieldName:="file", _
                    additionalFormData:=formData
    
    ' 处理响应
    If http.Inst.Status = 200 Then
        Debug.Print "上传成功!"
    End If
End Sub
```

### 链式调用上传

```vb
Sub ChainUpload()
    Dim http As New cHttpClient
    Dim result As cJson
    
    ' 上传并直接获取 JSON 响应
    Set result = http.UploadFile("https://api.example.com/upload", _
                                  "C:\Data\doc.docx", _
                                  "document") _
                         .ReturnJson()
    
    Debug.Print "文件ID: " & result.Item("id")
End Sub
```

### 传统方式上传（手动构建请求）

如需更多控制，仍可使用传统方式：

```vb
Sub ManualUpload()
    Dim http As New cHttpClient
    
    ' 读取文件到字节数组
    Dim Stream As New cToolsStream
    Dim fileBytes() As Byte
    Stream.LoadFileAsBinary "C:\data.pdf", fileBytes
    
    ' 设置 multipart 表单并发送
    http.SetRequestContentType(FormMultipart)
    http.RequestHeaders.Add "Content-Disposition", _
        "form-data; name=""file""; filename=""data.pdf"""
    
    http.SendPost "https://api.example.com/upload", _
                  StrConv(fileBytes, vbUnicode)
End Sub
```

---

## 文件下载

### 简单文件下载（推荐）

使用内置的 `DownloadFile` 方法一行代码完成下载：

```vb
Sub SimpleDownload()
    Dim http As New cHttpClient
    
    ' 下载文件（自动覆盖已存在文件）
    If http.DownloadFile("https://example.com/file.pdf", _
                          "C:\Downloads\file.pdf") Then
        Debug.Print "下载成功!"
    End If
    
    ' 不覆盖已存在的文件
    http.DownloadFile "https://example.com/file.pdf", _
                      "C:\Downloads\file.pdf", _
                      Overwrite:=False
End Sub
```

### 异步文件下载

下载大文件时使用异步方式避免阻塞：

```vb
Private WithEvents httpDownload As cHttpClient

Sub AsyncDownload()
    Set httpDownload = New cHttpClient
    
    ' 启动异步下载
    httpDownload.DownloadFileAsync "https://example.com/large-file.zip", _
                                   "C:\Downloads\large-file.zip"
    
    Debug.Print "下载已启动..."
End Sub

Private Sub httpDownload_OnResponseFinished()
    ' 下载完成后保存文件
    If httpDownload.FinishDownloadFile() Then
        Debug.Print "异步下载完成!"
    Else
        Debug.Print "下载失败: " & httpDownload.LastError
    End If
End Sub
```

### 批量下载

```vb
Sub BatchDownload()
    Dim http As New cHttpClient
    Dim files As Variant
    Dim i As Long
    
    files = Array( _
        Array("https://example.com/file1.pdf", "C:\Downloads\file1.pdf"), _
        Array("https://example.com/file2.pdf", "C:\Downloads\file2.pdf") _
    )
    
    For i = LBound(files) To UBound(files)
        On Error Resume Next
        http.DownloadFile files(i)(0), files(i)(1), Overwrite:=True
        
        If Err.Number = 0 Then
            Debug.Print "下载成功: " & files(i)(1)
        Else
            Debug.Print "下载失败: " & files(i)(0)
        End If
        On Error GoTo 0
    Next
End Sub
```

### 传统方式下载（手动保存）

如需更多控制，仍可使用传统方式：

```vb
Sub ManualDownload()
    Dim http As New cHttpClient
    
    http.SendGet("https://example.com/file.pdf")
    
    If http.Inst.Status = 200 Then
        ' 使用 cToolsStream 保存文件
        Dim Stream As New cToolsStream
        Dim fileData() As Byte
        fileData = http.ReturnBody()
        Stream.SaveFileAsBinary "C:\Downloads\file.pdf", fileData
        
        Debug.Print "文件已保存"
    End If
End Sub
```

## OAuth 认证流程

### 获取 Access Token

```vb
Function GetAccessToken() As String
    Dim http As New cHttpClient
    
    http.SetRequestContentType(FormUrlEncoded)
    http.RequestDataForm.Add "grant_type", "client_credentials"
    http.RequestDataForm.Add "client_id", "your_client_id"
    http.RequestDataForm.Add "client_secret", "your_client_secret"
    
    http.SendPost("https://oauth.example.com/token")
    
    If http.Inst.Status = 200 Then
        Dim json As cJson
        Set json = http.ReturnJson()
        GetAccessToken = json.Item("access_token")
    End If
End Function
```

### 使用 Token 访问 API

```vb
Sub CallApiWithToken()
    Dim http As New cHttpClient
    
    ' 设置认证头
    http.RequestHeaders.Add "Authorization", "Bearer " & GetAccessToken()
    
    http.SendGet("https://api.example.com/protected-resource")
    
    If http.Inst.Status = 200 Then
        ProcessData http.ReturnJson()
    ElseIf http.Inst.Status = 401 Then
        ' Token 过期，刷新 Token 后重试
        RefreshToken
        CallApiWithToken
    End If
End Sub
```

## 会话保持

### 使用同一实例保持会话

```vb
Sub SessionExample()
    Dim http As New cHttpClient
    
    ' 登录
    http.RequestDataForm.Add "username", "admin"
    http.RequestDataForm.Add "password", "123456"
    http.SendPost("https://api.example.com/login")
    
    ' WinHttp 会自动处理 Set-Cookie
    
    ' 后续请求保持会话
    http.SendGet("https://api.example.com/user/profile")
    Debug.Print http.ReturnText()
    
    ' 再次请求，会话仍然有效
    http.SendGet("https://api.example.com/user/settings")
    Debug.Print http.ReturnText()
End Sub
```

## 分页处理

### 自动遍历分页 API

```vb
Sub FetchAllPages()
    Dim http As New cHttpClient
    Dim allData As New Collection
    
    Dim page As Long
    page = 1
    
    Do
        http.RequestDataQuery.RemoveAll
        http.RequestDataQuery.Add "page", page
        http.RequestDataQuery.Add "limit", 100
        
        http.SendGet("https://api.example.com/items")
        
        If http.Inst.Status <> 200 Then Exit Do
        
        Dim json As cJson
        Set json = http.ReturnJson()
        
        Dim items As Object
        Set items = json.Item("data")
        
        Dim i As Long
        For i = 0 To items.Count - 1
            allData.Add items(i)
        Next
        
        ' 检查是否还有下一页
        If items.Count < 100 Then Exit Do
        page = page + 1
        
    Loop
    
    Debug.Print "共获取 " & allData.Count & " 条数据"
End Sub
```

## 代理设置

### 设置代理服务器

```vb
Sub WithProxy()
    Dim http As New cHttpClient
    
    ' WinHttp 代理设置（需要直接操作 Inst）
    http.Inst.SetProxy 2, "http://proxy.example.com:8080"
    
    ' 如果代理需要认证
    http.Inst.SetCredentials "proxy_user", "proxy_pass", _
        HTTPREQUEST_SETCREDENTIALS_FOR_PROXY
    
    http.SendGet("https://api.example.com/data")
End Sub
```

## 自定义超时

### 精细控制各阶段超时

```vb
Sub CustomTimeouts()
    Dim http As New cHttpClient
    
    ' 直接操作 WinHttpRequest 设置更细粒度的超时
    ' SetTimeouts(解析DNS超时, 连接超时, 发送超时, 接收超时)
    http.Inst.SetTimeouts 10000, 10000, 60000, 60000
    
    http.RequestTimeOut = 60  ' 同步等待超时
    http.SendGet("https://slow-api.example.com/data")
End Sub
```

## 响应内容处理

### 根据 Content-Type 自动处理

```vb
Sub AutoProcessResponse()
    Dim http As New cHttpClient
    
    http.SendGet("https://api.example.com/data")
    
    Dim contentType As String
    If http.ResponseHeaders.Exists("Content-Type") Then
        contentType = http.ResponseHeaders("Content-Type")
    End If
    
    Select Case True
        Case InStr(contentType, "application/json") > 0
            ProcessJson http.ReturnJson()
            
        Case InStr(contentType, "application/xml") > 0
            ProcessXml http.ReturnText()
            
        Case InStr(contentType, "text/html") > 0
            ProcessHtml http.ReturnText()
            
        Case InStr(contentType, "image/") > 0
            SaveImage http.ReturnBody()
            
        Case Else
            ProcessRaw http.ReturnBody()
    End Select
End Sub
```

## 重试机制

### 自动重试失败的请求

```vb
Function RetryRequest(url As String, maxRetries As Long) As cJson
    Dim http As New cHttpClient
    Dim attempt As Long
    
    For attempt = 1 To maxRetries
        On Error Resume Next
        
        http.SendGet(url)
        
        If Err.Number = 0 And http.Inst.Status < 400 Then
            Set RetryRequest = http.ReturnJson()
            Exit Function
        End If
        
        On Error GoTo 0
        
        Debug.Print "请求失败，第 " & attempt & " 次重试..."
        Sleep 1000 * attempt  ' 指数退避
    Next
    
    Err.Raise vbObjectError + 1, , "请求在 " & maxRetries & " 次尝试后仍然失败"
End Function
```
