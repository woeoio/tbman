# API Reference

## Class Properties

### Request Related

| Property | Type | Description | Default |
|----------|------|-------------|---------|
| `RequestData()` | Byte() | Request body byte array | - |
| `RequestDataBody` | Dictionary | Body data dictionary | - |
| `RequestDataForm` | Dictionary | Form data | - |
| `RequestDataJson` | cJson | JSON data | - |
| `RequestDataQuery` | Dictionary | URL query parameters | - |
| `RequestHeaders` | Dictionary | Request headers | - |
| `RequestContentType` | String | Content-Type | - |
| `RequestChartSet` | String | Encoding | "utf-8" |
| `RequestTimeOut` | Long | Timeout (seconds) | 5 |

### Response Related

| Property | Type | Description |
|----------|------|-------------|
| `ResponseRaw` | Variant | Raw response content |
| `ResponseHeaders` | Dictionary | Response headers |
| `Cookies` | Dictionary | Cookie dictionary |

### Control Options

| Property | Type | Description | Default |
|----------|------|-------------|---------|
| `EnableRedirects` | Boolean | Auto redirect | True |
| `MaxRedirects` | Long | Max redirects | 10 |
| `DebugStart` | Boolean | Debug mode | False |
| `DebugInfo` | cJson | Debug info | - |
| `LastError` | String | Last error message | "" |

### Internal Objects

| Property | Type | Description |
|----------|------|-------------|
| `Inst` | WinHttp.WinHttpRequest | WinHTTP instance |

---

## Class Methods

### HTTP Request Methods

#### SendGet
```vb
Public Function SendGet(ByVal url As String, Optional Body As String) As cHttpClient
```
Send GET request.

#### SendPost
```vb
Public Function SendPost(ByVal url As String, Optional Body As String) As cHttpClient
```
Send POST request.

#### SendPut
```vb
Public Function SendPut(ByVal url As String, Optional Body As String) As cHttpClient
```
Send PUT request.

#### SendDelete
```vb
Public Function SendDelete(ByVal url As String, Optional Body As String) As cHttpClient
```
Send DELETE request.

#### SendOptions
```vb
Public Function SendOptions(ByVal url As String, Optional Body As String) As cHttpClient
```
Send OPTIONS request.

#### Send
```vb
Public Function Send(Method As EnumRequestMethod, ByVal Url As String, Optional Body As String) As cHttpClient
```
Generic request method.

#### Fetch
```vb
Public Function Fetch(Method As EnumRequestMethod, ByVal url As String, Optional Body As String) As cHttpClient
```
Alias for Send.

---

### Configuration Methods

#### SetRequestContentType
```vb
Public Function SetRequestContentType(ReqType As EnumRequestContentType, Optional ContentType As String) As String
```
Set request Content-Type.

| Parameter | Description |
|-----------|-------------|
| `ReqType` | Enum value 0-5 |
| `ContentType` | Custom Content-Type (optional) |

**Return:** Actual Content-Type string set

#### MapRequestContentType
```vb
Public Function MapRequestContentType(ReqType As EnumRequestContentType, Optional ContentType As String) As String
```
Alias for SetRequestContentType.

#### Async
```vb
Public Function Async(Bool As Boolean) As cHttpClient
```
Set sync/async mode.

#### SetCookies
```vb
Public Function SetCookies(ByVal Value As String) As cHttpClient
```
Set cookie.

#### AllowRedirects
```vb
Public Function AllowRedirects(Bool As Boolean) As cHttpClient
```
Set whether to allow auto redirect.

#### ResetRedirectCount
```vb
Public Function ResetRedirectCount() As cHttpClient
```
Reset redirect counter.

---

### Response Handling Methods

#### ReturnText
```vb
Public Function ReturnText(Optional IsUtf8 As Boolean = True, Optional IsConvert As Boolean) As String
```
Get response text.

| Parameter | Description | Default |
|-----------|-------------|---------|
| `IsUtf8` | Use UTF-8 decoding | True |
| `IsConvert` | Use StrConv conversion | False |

#### ReturnJson
```vb
Public Function ReturnJson(Optional IsUtf8 As Boolean = True, Optional IsConvert As Boolean) As cJson
```
Get response JSON object.

#### ReturnBody
```vb
Public Function ReturnBody() As Byte()
```
Get response byte array.

#### ReturnStream
```vb
Public Function ReturnStream() As Variant
```
Get response stream object.

---

### Redirect Methods

#### GetRedirectUrl
```vb
Public Function GetRedirectUrl() As String
```
Get redirect target URL (from Location header).

**Return:** Redirect URL, empty string if not exists

#### FollowRedirect
```vb
Public Function FollowRedirect(Optional RedirectMethod As EnumRequestMethod = 0) As cHttpClient
```
Manually follow redirect.

**Parameters:**

| Parameter | Type | Description | Default |
|-----------|------|-------------|---------|
| `RedirectMethod` | EnumRequestMethod | Request method for 307/308 redirect | `0` (auto handle) |

Behavior based on status code:
- 301/302/303: Convert to GET
- 307/308: Default GET, can specify other method via `RedirectMethod` parameter
- Others: Default GET

**Example:**
```vb
' Default behavior: 307/308 use GET
http.FollowRedirect

' Specify 307/308 use POST (preserve method)
http.FollowRedirect(ReqPost)

' Specify use PUT
http.FollowRedirect(ReqPut)
```

---

### File Operation Methods

#### DownloadFile
```vb
Public Function DownloadFile(ByVal url As String, ByVal savePath As String, Optional Overwrite As Boolean = True) As Boolean
```
Synchronously download file and save to specified path.

| Parameter | Type | Description | Default |
|-----------|------|-------------|---------|
| `url` | String | File URL | Required |
| `savePath` | String | Local save path | Required |
| `Overwrite` | Boolean | Overwrite existing file | `True` |

**Return:** Success returns `True`, failure throws error

#### DownloadFileAsync
```vb
Public Function DownloadFileAsync(ByVal url As String, ByVal savePath As String, Optional Overwrite As Boolean = True) As cHttpClient
```
Start async file download.

**Note:** Need to call `FinishDownloadFile` in `OnResponseFinished` event to complete save.

#### FinishDownloadFile
```vb
Public Function FinishDownloadFile() As Boolean
```
Complete async download, save file. Call in `OnResponseFinished` event.

#### UploadFile
```vb
Public Function UploadFile(ByVal url As String, ByVal filePath As String, Optional fieldName As String = "file", Optional additionalFormData As Scripting.Dictionary = Nothing) As cHttpClient
```
Upload file to server (using multipart/form-data format).

| Parameter | Type | Description | Default |
|-----------|------|-------------|---------|
| `url` | String | Upload URL | Required |
| `filePath` | String | Local file path | Required |
| `fieldName` | String | Form field name | `"file"` |
| `additionalFormData` | Dictionary | Additional form data | `Nothing` |

**Return:** Returns `cHttpClient` instance, supports method chaining

#### UploadFileSimple
```vb
Public Function UploadFileSimple(ByVal url As String, ByVal filePath As String) As Boolean
```
Simplified file upload, returns boolean for success/failure.

### Utility Methods

#### ShowPage
```vb
Public Sub ShowPage(url As String)
```
Open URL with default system browser.

---

## Enum Types

### EnumRequestMethod

| Value | Name | Description |
|-------|------|-------------|
| 0 | ReqGet | GET request |
| 1 | ReqPost | POST request |
| 2 | ReqPut | PUT request |
| 3 | ReqDelete | DELETE request |
| 4 | ReqOptions | OPTIONS request |
| 5 | ReqPatch | PATCH request |

### EnumRequestContentType

| Value | Name | Content-Type |
|-------|------|--------------|
| 0 | - | Empty |
| 1 | Json | application/json |
| 2 | FormUrlEncoded | application/x-www-form-urlencoded |
| 3 | FormMultipart | multipart/form-data |
| 4 | TextPlain | text/plain |
| 5 | TextHtml | text/html |

---

## Events

### OnResponseFinished

```vb
Public Event OnResponseFinished()
```
Triggered when async request completes.

**Usage Example:**
```vb
Private WithEvents http As cHttpClient

Private Sub http_OnResponseFinished()
    Debug.Print "Request complete"
    Debug.Print http.ReturnText()
End Sub

Sub StartAsyncRequest()
    Set http = New cHttpClient
    http.Async(True).SendGet("https://api.example.com/data")
End Sub
```
