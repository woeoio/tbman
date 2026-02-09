# Method Chaining

Method chaining is a core design feature of `cHttpClient`, achieved by returning the object instance itself for a fluent API. This design allows you to **complete an entire request in one statement**, or **organize code in segments**.

---

## Two Calling Styles

### Style 1: One-liner Request

Suitable for simple request scenarios, concise and efficient code:

```vb
' Simplest GET request
MsgBox TBMAN.HttpClient.Send(ReqGet, "http://a-vi.com/tbman/hello/woeoio").ReturnText()

' One-liner with configuration
MsgBox VBMAN.HttpClient.Fetch(ReqGet, "http://a-vi.com/hello/VBMAN").ReturnJson().Encode(, 2, True)
```

**Advantages:**
- Extremely concise code
- Suitable for one-time, simple requests
- No variable declaration needed

### Style 2: Segmented Request (Recommended)

Suitable for complex request scenarios, clear code structure:

```vb
With New cHttpClient
    ' 1. Configuration phase
    .RequestHeaders.Add "Accept", "application/json"
    .RequestHeaders.Add "AccessToken", "5021a00d60db149fa8931f66e2f9d854"
    .SetRequestContentType JsonString
    .DebugStart = True

    ' 2. Build request body
    Dim Body As String
    With New cJson
        .Item("username") = "dengwei"
        .Item("password") = "123456"
        Body = .Encode()
    End With

    ' 3. Send request
    .SendPost "https://api.vb6.pro/tbman/login?ret=0", Body

    ' 4. Handle response
    With .ReturnJson()
        If .Root("code") = 0 Then
            MsgBox("wellcome to www.vb6.pro")
        Else
            MsgBox(.Root("msg"))
        End If
    End With

    ' 5. Output debug info
    Debug.Print .DebugInfo.Encode(, 2, True)
End With
```

**Advantages:**
- Clear logical layering
- Easy to maintain and debug
- Suitable for complex business scenarios

---

## Method Chaining List

The following methods all return `cHttpClient` instance, supporting method chaining:

| Method | Description | Return Value |
|--------|-------------|--------------|
| `SetRequestContentType()` | Set Content-Type | cHttpClient |
| `MapRequestContentType()` | Alias for SetRequestContentType | cHttpClient |
| `Async()` | Set sync/async mode | cHttpClient |
| `SetCookies()` | Set cookie | cHttpClient |
| `AllowRedirects()` | Set whether to allow redirect | cHttpClient |
| `ResetRedirectCount()` | Reset redirect counter | cHttpClient |
| `SendPost()` | Send POST request | cHttpClient |
| `SendGet()` | Send GET request | cHttpClient |
| `SendPut()` | Send PUT request | cHttpClient |
| `SendDelete()` | Send DELETE request | cHttpClient |
| `SendOptions()` | Send OPTIONS request | cHttpClient |
| `Send()` | Generic request method | cHttpClient |
| `Fetch()` | Alias for Send | cHttpClient |
| `FollowRedirect()` | Manually follow redirect | cHttpClient |

**Note:** `ReturnText()`, `ReturnJson()`, `ReturnBody()` and other response handling methods **do not** return cHttpClient instance, they should be placed at the end of method chaining.

---

## Practical Examples

### Example 1: One-liner POST Request

```vb
' Send POST and directly get result
Dim result As String
result = VBMAN.HttpClient.SetRequestContentType(JsonString).SendPost("https://api.example.com/login", Body).ReturnText()
```

### Example 2: Segmented JSON Request Building

```vb
With New cHttpClient
    ' Configure request
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

    ' Send and handle
    .SetRequestContentType(JsonString).SendPost("https://api.example.com/order")

    With .ReturnJson()
        MsgBox .Root("orderId")
    End With
End With
```

### Example 3: Mixed Style - Chaining + Segmented

```vb
' Chain configuration, segmented logic
With VBMAN.HttpClient
    ' Chain config (set multiple properties in one line)
    .SetRequestContentType(JsonString).DebugStart = True

    ' Segmented build complex request body
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

    ' Chain send request and get result
    Text1.Text = .Fetch(ReqPost, Url, PostBody).ReturnJson().Encode(, 2, True)
End With
```

---

## Best Practices

### 1. Simple Requests with One-liner

```vb
' ✅ Recommended: Simple request direct chaining
Dim json As cJson
Set json = VBMAN.HttpClient.Fetch(ReqGet, "https://api.example.com/data").ReturnJson()
```

### 2. Complex Requests Segmented

```vb
' ✅ Recommended: Complex request segmented handling
With New cHttpClient
    ' Configure request headers
    .RequestHeaders.Add "X-API-Key", apiKey
    .RequestHeaders.Add "Authorization", "Bearer " & token

    ' Build request body
    ' ... complex data building logic ...

    ' Send request
    .SendPost(url, body)

    ' Handle response
    ' ... complex response handling logic ...
End With
```

### 3. Avoid Over-chaining

```vb
' ❌ Not recommended: Too long one line, hard to read and maintain
MsgBox VBMAN.HttpClient.SetRequestContentType(JsonString).SetCookies("session=abc").Async(False).AllowRedirects(True).SendPost(url, body).ReturnJson().Item("data").Item("name")

' ✅ Recommended: Appropriate segmentation, maintain readability
With VBMAN.HttpClient
    .SetRequestContentType(JsonString)
    .SetCookies("session=abc")
    .SendPost url, body
    MsgBox .ReturnJson().Item("data").Item("name")
End With
```

### 4. Error Handling

```vb
' ✅ Recommended: Handle errors within With block
With New cHttpClient
    On Error GoTo ErrorHandler

    .SendGet "https://api.example.com/data"
    Debug.Print .ReturnText()

    Exit Sub

ErrorHandler:
    Debug.Print "Request failed: " & Err.Description
    Debug.Print "Debug info: " & .DebugInfo.Encode(, 2, True)
End With
```

---

## Essence of Method Chaining

Each chained method has this line at the end:

```vb
Public Function Xxx() As cHttpClient
    ' ... logic code ...
    Set Xxx = Me  ' Return self instance
End Function
```

This is why you can write:

```vb
.SetRequestContentType(JsonString).SendPost(url, body).ReturnText()
'     ↑ Returns cHttpClient        ↑ Returns cHttpClient   ↑ Returns String
```

---

## Compare with Traditional Style

### Traditional Style (Verbose)

```vb
Dim http As cHttpClient
Set http = New cHttpClient

http.SetRequestContentType(JsonString)
http.SendPost url, body

Dim json As cJson
Set json = http.ReturnJson()

MsgBox json.Item("name")
```

### Chaining Style (Concise)

```vb
With New cHttpClient
    MsgBox .SetRequestContentType(JsonString).SendPost(url, body).ReturnJson().Item("name")
End With
```

---

## Summary

| Scenario | Recommended Style | Example |
|----------|-------------------|---------|
| Simple GET request | One-liner | `MsgBox http.SendGet(url).ReturnText()` |
| Request with parameters | Segmented | `With http: .AddHeader: .SendGet: End With` |
| Complex POST | Segmented | `With http: config-send-handle: End With` |
| Needs error handling | Segmented | `With http: On Error: request: End With` |
| Quick test | One-liner | `Debug.Print http.Fetch(...).ReturnText()` |

The core value of method chaining is **flexibility**—keeping simple tasks concise while keeping complex tasks clear. Choose the appropriate writing style based on the actual scenario.
