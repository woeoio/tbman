# 链式调用

链式调用是 `cHttpClient` 的核心设计特性，通过返回对象自身实例实现流畅的 API 接口。这种设计让你可以用**一句话完成整个请求**，也可以**分段式地组织代码**。

---

## 两种调用风格

### 风格一：一句话完成请求

适合简单的请求场景，代码简洁高效：

```vb
' 最简单的 GET 请求
MsgBox TBMAN.HttpClient.Send(ReqGet, "http://a-vi.com/tbman/hello/woeoio").ReturnText()

' 带配置的一行代码
MsgBox VBMAN.HttpClient.Fetch(ReqGet, "http://a-vi.com/hello/VBMAN").ReturnJson().Encode(, 2, True)
```

**优点：**
- 代码极度简洁
- 适合一次性、简单的请求
- 不需要声明变量

### 风格二：分段式请求（推荐）

适合复杂的请求场景，代码结构清晰：

```vb
With New cHttpClient
    ' 1. 配置阶段
    .RequestHeaders.Add "Accept", "application/json"
    .RequestHeaders.Add "AccessToken", "5021a00d60db149fa8931f66e2f9d854"
    .SetRequestContentType JsonString
    .DebugStart = True
    
    ' 2. 构造请求体
    Dim Body As String
    With New cJson
        .Item("username") = "dengwei"
        .Item("password") = "123456"
        Body = .Encode()
    End With
    
    ' 3. 发送请求
    .SendPost "https://api.vb6.pro/tbman/login?ret=0", Body
    
    ' 4. 处理响应
    With .ReturnJson()
        If .Root("code") = 0 Then
            MsgBox("wellcome to www.vb6.pro")
        Else
            MsgBox(.Root("msg"))
        End If
    End With
    
    ' 5. 输出调试信息
    Debug.Print .DebugInfo.Encode(, 2, True)
End With
```

**优点：**
- 逻辑分层清晰
- 易于维护和调试
- 适合复杂业务场景

---

## 链式调用方法清单

以下方法都返回 `cHttpClient` 实例，支持链式调用：

| 方法 | 说明 | 返回值 |
|------|------|--------|
| `SetRequestContentType()` | 设置 Content-Type | cHttpClient |
| `MapRequestContentType()` | SetRequestContentType 的别名 | cHttpClient |
| `Async()` | 设置同步/异步模式 | cHttpClient |
| `SetCookies()` | 设置 Cookie | cHttpClient |
| `AllowRedirects()` | 设置是否允许重定向 | cHttpClient |
| `ResetRedirectCount()` | 重置重定向计数 | cHttpClient |
| `SendPost()` | 发送 POST 请求 | cHttpClient |
| `SendGet()` | 发送 GET 请求 | cHttpClient |
| `SendPut()` | 发送 PUT 请求 | cHttpClient |
| `SendDelete()` | 发送 DELETE 请求 | cHttpClient |
| `SendOptions()` | 发送 OPTIONS 请求 | cHttpClient |
| `Send()` | 通用请求方法 | cHttpClient |
| `Fetch()` | Send 的别名 | cHttpClient |
| `FollowRedirect()` | 手动跟随重定向 | cHttpClient |

**注意：** `ReturnText()`、`ReturnJson()`、`ReturnBody()` 等响应处理方法**不返回** cHttpClient 实例，它们应该放在链式调用的最后。

---

## 实战示例

### 示例 1：一句话 POST 请求

```vb
' 发送 POST 并直接获取结果
Dim result As String
result = VBMAN.HttpClient.SetRequestContentType(JsonString).SendPost("https://api.example.com/login", Body).ReturnText()
```

### 示例 2：分段式构建 JSON 请求

```vb
With New cHttpClient
    ' 配置请求
    With .RequestDataJson
        With .NewItem("consignee")
            With .NewItem("address")
                .Item("city") = "Paderborn"
                .Item("countryCode") = "DE"
            End With
        End With
        With .NewItems("lines")
            With .NewItem()
                .Item("content") = "furniture"
                .Item("unitWeight") = 200
            End With
        End With
    End With
    
    ' 发送并处理
    .SetRequestContentType(JsonString).SendPost("https://api.example.com/order")
    
    With .ReturnJson()
        MsgBox .Root("orderId")
    End With
End With
```

### 示例 3：混合风格 - 链式 + 分段

```vb
' 链式设置配置，分段处理逻辑
With VBMAN.HttpClient
    ' 链式配置（一行设置多个属性）
    .SetRequestContentType(JsonString).DebugStart = True
    
    ' 分段构造复杂请求体
    Dim PostBody As String
    With New cJson
        .Item("sysStuffCode") = "TEST001"
        With .NewItems("detailList")
            Dim i As Long
            For i = 0 To 3
                With .NewItem()
                    .Item("test") = 123
                    .Item("time") = Now()
                End With
            Next
        End With
        PostBody = .Encode()
    End With
    
    ' 链式发送请求并获取结果
    Text1.Text = .Fetch(ReqPost, Url, PostBody).ReturnJson().Encode(, 2, True)
End With
```

---

## 最佳实践

### 1. 简单请求用一句话

```vb
' ✅ 推荐：简单请求直接链式调用
Dim json As cJson
Set json = VBMAN.HttpClient.Fetch(ReqGet, "https://api.example.com/data").ReturnJson()
```

### 2. 复杂请求用分段式

```vb
' ✅ 推荐：复杂请求分段处理
With New cHttpClient
    ' 配置请求头
    .RequestHeaders.Add "X-API-Key", apiKey
    .RequestHeaders.Add "Authorization", "Bearer " & token
    
    ' 构造请求体
    ' ... 复杂的数据构建逻辑 ...
    
    ' 发送请求
    .SendPost(url, body)
    
    ' 处理响应
    ' ... 复杂的响应处理逻辑 ...
End With
```

### 3. 避免过度链式

```vb
' ❌ 不推荐：过长的一行代码，难以阅读和维护
MsgBox VBMAN.HttpClient.SetRequestContentType(JsonString).SetCookies("session=abc").Async(False).AllowRedirects(True).SendPost(url, body).ReturnJson().Item("data").Item("name")

' ✅ 推荐：适当分段，保持可读性
With VBMAN.HttpClient
    .SetRequestContentType(JsonString)
    .SetCookies("session=abc")
    .SendPost url, body
    MsgBox .ReturnJson().Item("data").Item("name")
End With
```

### 4. 错误处理

```vb
' ✅ 推荐：在 With 块内处理错误
With New cHttpClient
    On Error GoTo ErrorHandler
    
    .SendGet "https://api.example.com/data"
    Debug.Print .ReturnText()
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "请求失败: " & Err.Description
    Debug.Print "调试信息: " & .DebugInfo.Encode(, 2, True)
End With
```

---

## 链式调用的本质

每个链式方法最后都有这行代码：

```vb
Public Function Xxx() As cHttpClient
    ' ... 逻辑代码 ...
    Set Xxx = Me  ' 返回自身实例
End Function
```

这就是为什么可以这样写：

```vb
.SetRequestContentType(JsonString).SendPost(url, body).ReturnText()
'     ↑ 返回 cHttpClient              ↑ 返回 cHttpClient      ↑ 返回 String
```

---

## 对比传统写法

### 传统写法（繁琐）

```vb
Dim http As cHttpClient
Set http = New cHttpClient

http.SetRequestContentType(JsonString)
http.SendPost url, body

Dim json As cJson
Set json = http.ReturnJson()

MsgBox json.Item("name")
```

### 链式写法（简洁）

```vb
With New cHttpClient
    MsgBox .SetRequestContentType(JsonString).SendPost(url, body).ReturnJson().Item("name")
End With
```

---

## 总结

| 场景 | 推荐风格 | 示例 |
|------|----------|------|
| 简单 GET 请求 | 一句话 | `MsgBox http.SendGet(url).ReturnText()` |
| 带参数的请求 | 分段式 | `With http: .AddHeader: .SendGet: End With` |
| 复杂 POST | 分段式 | `With http: 配置-发送-处理: End With` |
| 需要错误处理 | 分段式 | `With http: On Error: 请求: End With` |
| 快速测试 | 一句话 | `Debug.Print http.Fetch(...).ReturnText()` |

链式调用的核心价值在于**灵活性**——既能让简单任务保持简洁，又能让复杂任务保持清晰。根据实际场景选择合适的书写风格即可。
