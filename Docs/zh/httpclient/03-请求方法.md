# 请求方法

## 支持的 HTTP 方法

### GET 请求

```vb
Public Function SendGet(ByVal url As String, Optional Body As String) As cHttpClient
```

**示例：**
```vb
http.SendGet("https://api.example.com/users")
http.SendGet("https://api.example.com/users", "id=123")  ' 带请求体
```

### POST 请求

```vb
Public Function SendPost(ByVal url As String, Optional Body As String) As cHttpClient
```

**示例：**
```vb
' 直接传 Body
http.SendPost("https://api.example.com/users", "name=John&age=30")

' 使用表单数据
http.RequestDataForm.Add "name", "John"
http.RequestDataForm.Add "age", "30"
http.SendPost("https://api.example.com/users")

' 使用 JSON
http.SetRequestContentType(Json)
http.RequestDataJson.Add "name", "John"
http.SendPost("https://api.example.com/users")
```

### PUT 请求

```vb
Public Function SendPut(ByVal url As String, Optional Body As String) As cHttpClient
```

**示例：**
```vb
http.SetRequestContentType(Json)
http.RequestDataJson.Add "name", "Updated Name"
http.SendPut("https://api.example.com/users/1")
```

### DELETE 请求

```vb
Public Function SendDelete(ByVal url As String, Optional Body As String) As cHttpClient
```

**示例：**
```vb
http.SendDelete("https://api.example.com/users/1")
```

### OPTIONS 请求

```vb
Public Function SendOptions(ByVal url As String, Optional Body As String) As cHttpClient
```

**示例：**
```vb
http.SendOptions("https://api.example.com/users")
' 查看 ResponseHeaders 中的 Allow 字段获取支持的方法
```

### 通用请求方法

```vb
Public Function Send(Method As EnumRequestMethod, ByVal Url As String, Optional Body As String) As cHttpClient
```

**示例：**
```vb
http.Send(ReqPatch, "https://api.example.com/users/1", "{\"name\":\"test\"}")
```

## 方法别名

`Fetch` 是 `Send` 的别名：

```vb
http.Fetch(ReqGet, "https://api.example.com/data")  ' 等同于 Send
```

## 添加 URL 查询参数

```vb
' 使用 RequestDataQuery 添加查询参数
http.RequestDataQuery.Add "page", "1"
http.RequestDataQuery.Add "limit", "10"
http.SendGet("https://api.example.com/users")
' 最终 URL: https://api.example.com/users?page=1&limit=10
```
