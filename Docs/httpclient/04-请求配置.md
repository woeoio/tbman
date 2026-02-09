# 请求配置

## Content-Type 设置

### 使用枚举设置

```vb
Public Function SetRequestContentType(ReqType As EnumRequestContentType, Optional ContentType As String) As String
```

| 枚举值 | Content-Type |
|--------|--------------|
| 0 | 空（不设置） |
| 1 | `application/json` |
| 2 | `application/x-www-form-urlencoded` |
| 3 | `multipart/form-data` |
| 4 | `text/plain` |
| 5 | `text/html` |

**示例：**
```vb
http.SetRequestContentType(Json)           ' application/json
http.SetRequestContentType(FormUrlEncoded) ' application/x-www-form-urlencoded
```

### 别名方法

```vb
http.MapRequestContentType(Json)  ' 等同于 SetRequestContentType
```

## 请求头设置

```vb
' 添加单个请求头
http.RequestHeaders.Add "Authorization", "Bearer token123"
http.RequestHeaders.Add "X-Custom-Header", "value"

' 设置 Cookie
http.SetCookies("session=abc123; user=john")
```

## 超时设置

```vb
Public RequestTimeOut As Long  ' 单位：秒，默认 5 秒
```

**示例：**
```vb
http.RequestTimeOut = 30  ' 30 秒超时
```

## 编码设置

```vb
Public RequestChartSet As String  ' 默认 "utf-8"
```

**示例：**
```vb
http.RequestChartSet = "utf-8"   ' UTF-8 编码
http.RequestChartSet = "gb2312"  ' GB2312 编码
```

## 异步模式

```vb
Public Function Async(Bool As Boolean) As cHttpClient
```

**示例：**
```vb
' 同步模式（默认）
http.Async(False).SendGet("https://api.example.com/data")

' 异步模式
http.Async(True).SendGet("https://api.example.com/data")
' 需要处理 OnResponseFinished 事件
```

## 数据字典

### 表单数据 (RequestDataForm)

```vb
http.RequestDataForm.Add "key1", "value1"
http.RequestDataForm.Add "key2", "value2"
http.SendPost("https://api.example.com/submit")
' 自动编码为：key1=value1&key2=value2
```

### JSON 数据 (RequestDataJson)

```vb
http.SetRequestContentType(Json)
http.RequestDataJson.Add "name", "John"
http.RequestDataJson.Add "age", 30
http.RequestDataJson.Add "items", Array("a", "b", "c")
http.SendPost("https://api.example.com/users")
```

### 查询参数 (RequestDataQuery)

```vb
http.RequestDataQuery.Add "page", 1
http.RequestDataQuery.Add "limit", 20
http.SendGet("https://api.example.com/users")
' URL 自动附加：?page=1&limit=20
```

### Body 数据 (RequestDataBody)

用于直接传递字节数组数据：

```vb
http.RequestDataBody.Add "raw", byteArray
```

## SSL/TLS 设置

自动忽略 SSL 证书错误：

```vb
' 内部自动设置，无需手动调用
Inst.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = &H3300
```
