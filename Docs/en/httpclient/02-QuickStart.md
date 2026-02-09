# Quick Start

## Basic Usage

### 1. Create Instance

```vb
Dim http As New cHttpClient
```

### 2. Send GET Request

```vb
' Simplest GET request
http.SendGet("https://jsonplaceholder.typicode.com/posts/1")
Debug.Print http.ReturnText()
```

### 3. Send POST Request

```vb
' Form data
http.RequestDataForm.Add "username", "admin"
http.RequestDataForm.Add "password", "123456"
http.SendPost("https://api.example.com/login")
```

### 4. Handle JSON Data

```vb
' Send JSON
http.SetRequestContentType(Json)
http.RequestDataJson.Add "title", "Hello"
http.RequestDataJson.Add "body", "World"
http.SendPost("https://api.example.com/posts")

' Parse JSON response
Dim json As cJson
Set json = http.ReturnJson()
Debug.Print json.Item("id")
```

## Method Chaining

All configuration methods return `cHttpClient` instance, supporting method chaining:

```vb
http.SetCookies("session=abc123") _
    .SetRequestContentType(Json) _
    .Async(False) _
    .SendGet("https://api.example.com/data")
```

## Error Handling

```vb
On Error GoTo ErrorHandler

http.SendGet("https://api.example.com/data")
Debug.Print http.ReturnText()

Exit Sub

ErrorHandler:
    Debug.Print "Error: " & Err.Description
    ' Get detailed debug info
    If http.DebugStart Then
        Debug.Print http.DebugInfo.Encode()
    End If
```

## Complete Example

```vb
Sub TestHttpClient()
    Dim http As New cHttpClient

    ' Enable debug mode
    http.DebugStart = True

    ' Set request headers
    http.RequestHeaders.Add "X-API-Key", "secret123"
    http.SetCookies("session=abc")

    ' Set timeout (seconds)
    http.RequestTimeOut = 30

    ' Send request
    http.SendGet("https://httpbin.org/get")

    ' Output result
    Debug.Print "Status: " & http.Inst.Status
    Debug.Print "Response: " & http.ReturnText()

    ' Output debug info
    Debug.Print http.DebugInfo.Encode(, , True)
End Sub
```
