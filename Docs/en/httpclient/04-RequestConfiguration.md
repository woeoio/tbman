# Request Configuration

## Content-Type Setting

### Using Enums

```vb
Public Function SetRequestContentType(ReqType As EnumRequestContentType, Optional ContentType As String) As String
```

| Enum Value | Content-Type |
|------------|--------------|
| 0 | Empty (not set) |
| 1 | `application/json` |
| 2 | `application/x-www-form-urlencoded` |
| 3 | `multipart/form-data` |
| 4 | `text/plain` |
| 5 | `text/html` |

**Example:**
```vb
http.SetRequestContentType(Json)           ' application/json
http.SetRequestContentType(FormUrlEncoded) ' application/x-www-form-urlencoded
```

### Alias Method

```vb
http.MapRequestContentType(Json)  ' Same as SetRequestContentType
```

## Request Headers

```vb
' Add single header
http.RequestHeaders.Add "Authorization", "Bearer token123"
http.RequestHeaders.Add "X-Custom-Header", "value"

' Set Cookie
http.SetCookies("session=abc123; user=john")
```

## Timeout Setting

```vb
Public RequestTimeOut As Long  ' Unit: seconds, default 5 seconds
```

**Example:**
```vb
http.RequestTimeOut = 30  ' 30 second timeout
```

## Encoding Setting

```vb
Public RequestChartSet As String  ' Default "utf-8"
```

**Example:**
```vb
http.RequestChartSet = "utf-8"   ' UTF-8 encoding
http.RequestChartSet = "gb2312"  ' GB2312 encoding
```

## Async Mode

```vb
Public Function Async(Bool As Boolean) As cHttpClient
```

**Example:**
```vb
' Synchronous mode (default)
http.Async(False).SendGet("https://api.example.com/data")

' Asynchronous mode
http.Async(True).SendGet("https://api.example.com/data")
' Need to handle OnResponseFinished event
```

## Data Dictionaries

### Form Data (RequestDataForm)

```vb
http.RequestDataForm.Add "key1", "value1"
http.RequestDataForm.Add "key2", "value2"
http.SendPost("https://api.example.com/submit")
' Automatically encoded as: key1=value1&key2=value2
```

### JSON Data (RequestDataJson)

```vb
http.SetRequestContentType(Json)
http.RequestDataJson.Add "name", "John"
http.RequestDataJson.Add "age", 30
http.RequestDataJson.Add "items", Array("a", "b", "c")
http.SendPost("https://api.example.com/users")
```

### Query Parameters (RequestDataQuery)

```vb
http.RequestDataQuery.Add "page", 1
http.RequestDataQuery.Add "limit", 20
http.SendGet("https://api.example.com/users")
' URL automatically appends: ?page=1&limit=20
```

### Body Data (RequestDataBody)

Used for directly passing byte array data:

```vb
http.RequestDataBody.Add "raw", byteArray
```

## SSL/TLS Settings

Automatically ignores SSL certificate errors:

```vb
' Internally auto-set, no manual call needed
Inst.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = &H3300
```
