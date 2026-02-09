# API 参考

## 类属性

### 请求相关

| 属性 | 类型 | 说明 | 默认值 |
|------|------|------|--------|
| `RequestData()` | Byte() | 请求体字节数组 | - |
| `RequestDataBody` | Dictionary | Body 数据字典 | - |
| `RequestDataForm` | Dictionary | 表单数据 | - |
| `RequestDataJson` | cJson | JSON 数据 | - |
| `RequestDataQuery` | Dictionary | URL 查询参数 | - |
| `RequestHeaders` | Dictionary | 请求头 | - |
| `RequestContentType` | String | Content-Type | - |
| `RequestChartSet` | String | 编码 | "utf-8" |
| `RequestTimeOut` | Long | 超时时间（秒） | 5 |

### 响应相关

| 属性 | 类型 | 说明 |
|------|------|------|
| `ResponseRaw` | Variant | 原始响应内容 |
| `ResponseHeaders` | Dictionary | 响应头 |
| `Cookies` | Dictionary | Cookie 字典 |

### 控制选项

| 属性 | 类型 | 说明 | 默认值 |
|------|------|------|--------|
| `EnableRedirects` | Boolean | 自动重定向 | True |
| `MaxRedirects` | Long | 最大重定向次数 | 10 |
| `DebugStart` | Boolean | 调试模式 | False |
| `DebugInfo` | cJson | 调试信息 | - |
| `LastError` | String | 最后错误信息 | "" |

### 内部对象

| 属性 | 类型 | 说明 |
|------|------|------|
| `Inst` | WinHttp.WinHttpRequest | WinHTTP 实例 |

---

## 类方法

### HTTP 请求方法

#### SendGet
```vb
Public Function SendGet(ByVal url As String, Optional Body As String) As cHttpClient
```
发送 GET 请求。

#### SendPost
```vb
Public Function SendPost(ByVal url As String, Optional Body As String) As cHttpClient
```
发送 POST 请求。

#### SendPut
```vb
Public Function SendPut(ByVal url As String, Optional Body As String) As cHttpClient
```
发送 PUT 请求。

#### SendDelete
```vb
Public Function SendDelete(ByVal url As String, Optional Body As String) As cHttpClient
```
发送 DELETE 请求。

#### SendOptions
```vb
Public Function SendOptions(ByVal url As String, Optional Body As String) As cHttpClient
```
发送 OPTIONS 请求。

#### Send
```vb
Public Function Send(Method As EnumRequestMethod, ByVal Url As String, Optional Body As String) As cHttpClient
```
通用请求方法。

#### Fetch
```vb
Public Function Fetch(Method As EnumRequestMethod, ByVal url As String, Optional Body As String) As cHttpClient
```
Send 的别名。

---

### 配置方法

#### SetRequestContentType
```vb
Public Function SetRequestContentType(ReqType As EnumRequestContentType, Optional ContentType As String) As String
```
设置请求 Content-Type。

| 参数 | 说明 |
|------|------|
| `ReqType` | 0-5 的枚举值 |
| `ContentType` | 自定义 Content-Type（可选） |

**返回值：** 实际设置的 Content-Type 字符串

#### MapRequestContentType
```vb
Public Function MapRequestContentType(ReqType As EnumRequestContentType, Optional ContentType As String) As String
```
SetRequestContentType 的别名。

#### Async
```vb
Public Function Async(Bool As Boolean) As cHttpClient
```
设置同步/异步模式。

#### SetCookies
```vb
Public Function SetCookies(ByVal Value As String) As cHttpClient
```
设置 Cookie。

#### AllowRedirects
```vb
Public Function AllowRedirects(Bool As Boolean) As cHttpClient
```
设置是否允许自动重定向。

#### ResetRedirectCount
```vb
Public Function ResetRedirectCount() As cHttpClient
```
重置重定向计数器。

---

### 响应处理方法

#### ReturnText
```vb
Public Function ReturnText(Optional IsUtf8 As Boolean = True, Optional IsConvert As Boolean) As String
```
获取响应文本。

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `IsUtf8` | 使用 UTF-8 解码 | True |
| `IsConvert` | 使用 StrConv 转换 | False |

#### ReturnJson
```vb
Public Function ReturnJson(Optional IsUtf8 As Boolean = True, Optional IsConvert As Boolean) As cJson
```
获取响应 JSON 对象。

#### ReturnBody
```vb
Public Function ReturnBody() As Byte()
```
获取响应字节数组。

#### ReturnStream
```vb
Public Function ReturnStream() As Variant
```
获取响应流对象。

---

### 重定向方法

#### GetRedirectUrl
```vb
Public Function GetRedirectUrl() As String
```
获取重定向目标 URL（从 Location 头）。

**返回值：** 重定向 URL，如果不存在返回空字符串

#### FollowRedirect
```vb
Public Function FollowRedirect(Optional RedirectMethod As EnumRequestMethod = 0) As cHttpClient
```
手动跟随重定向。

**参数：**

| 参数 | 类型 | 说明 | 默认值 |
|------|------|------|--------|
| `RedirectMethod` | EnumRequestMethod | 指定 307/308 重定向时使用的请求方法 | `0`（自动处理） |

根据状态码决定行为：
- 301/302/303：转为 GET 请求
- 307/308：默认使用 GET，可通过 `RedirectMethod` 参数指定其他方法
- 其他：默认使用 GET

**示例：**
```vb
' 默认行为：307/308 使用 GET
http.FollowRedirect

' 指定 307/308 使用 POST（保持原方法）
http.FollowRedirect(ReqPost)

' 指定使用 PUT
http.FollowRedirect(ReqPut)
```

---

### 文件操作方法

#### DownloadFile
```vb
Public Function DownloadFile(ByVal url As String, ByVal savePath As String, Optional Overwrite As Boolean = True) As Boolean
```
同步下载文件并保存到指定路径。

| 参数 | 类型 | 说明 | 默认值 |
|------|------|------|--------|
| `url` | String | 文件 URL 地址 | 必需 |
| `savePath` | String | 本地保存路径 | 必需 |
| `Overwrite` | Boolean | 是否覆盖已存在文件 | `True` |

**返回值：** 成功返回 `True`，失败抛出错误

#### DownloadFileAsync
```vb
Public Function DownloadFileAsync(ByVal url As String, ByVal savePath As String, Optional Overwrite As Boolean = True) As cHttpClient
```
启动异步文件下载。

**注意：** 需要在 `OnResponseFinished` 事件中调用 `FinishDownloadFile` 完成保存。

#### FinishDownloadFile
```vb
Public Function FinishDownloadFile() As Boolean
```
完成异步下载，保存文件。在 `OnResponseFinished` 事件中调用。

#### UploadFile
```vb
Public Function UploadFile(ByVal url As String, ByVal filePath As String, Optional fieldName As String = "file", Optional additionalFormData As Scripting.Dictionary = Nothing) As cHttpClient
```
上传文件到服务器（使用 multipart/form-data 格式）。

| 参数 | 类型 | 说明 | 默认值 |
|------|------|------|--------|
| `url` | String | 上传地址 | 必需 |
| `filePath` | String | 本地文件路径 | 必需 |
| `fieldName` | String | 表单字段名 | `"file"` |
| `additionalFormData` | Dictionary | 额外的表单数据 | `Nothing` |

**返回值：** 返回 `cHttpClient` 实例，支持链式调用

#### UploadFileSimple
```vb
Public Function UploadFileSimple(ByVal url As String, ByVal filePath As String) As Boolean
```
简化的文件上传，返回布尔值表示成功/失败。

### 工具方法

#### ShowPage
```vb
Public Sub ShowPage(url As String)
```
使用系统默认浏览器打开 URL。

---

## 枚举类型

### EnumRequestMethod

| 值 | 名称 | 说明 |
|----|------|------|
| 0 | ReqGet | GET 请求 |
| 1 | ReqPost | POST 请求 |
| 2 | ReqPut | PUT 请求 |
| 3 | ReqDelete | DELETE 请求 |
| 4 | ReqOptions | OPTIONS 请求 |
| 5 | ReqPatch | PATCH 请求 |

### EnumRequestContentType

| 值 | 名称 | Content-Type |
|----|------|--------------|
| 0 | - | 空 |
| 1 | Json | application/json |
| 2 | FormUrlEncoded | application/x-www-form-urlencoded |
| 3 | FormMultipart | multipart/form-data |
| 4 | TextPlain | text/plain |
| 5 | TextHtml | text/html |

---

## 事件

### OnResponseFinished

```vb
Public Event OnResponseFinished()
```
异步请求完成时触发。

**使用示例：**
```vb
Private WithEvents http As cHttpClient

Private Sub http_OnResponseFinished()
    Debug.Print "请求完成"
    Debug.Print http.ReturnText()
End Sub

Sub StartAsyncRequest()
    Set http = New cHttpClient
    http.Async(True).SendGet("https://api.example.com/data")
End Sub
```
