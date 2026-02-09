# Redirect Handling

## Auto Redirect (Default)

WinHttpRequest automatically follows 3xx redirects by default:

```vb
Dim http As New cHttpClient
http.SendGet("https://bit.ly/3xxxxx")  ' Auto follow 301/302
Debug.Print http.ReturnText()          ' Get final page content
```

## Disable Auto Redirect

### Disable Redirect

```vb
http.AllowRedirects(False).SendGet("https://example.com/redirect")
Debug.Print http.Inst.Status           ' 302
Debug.Print http.GetRedirectUrl()      ' Get redirect URL from Location header
```

### AllowRedirects Method

```vb
Public Function AllowRedirects(Bool As Boolean) As cHttpClient
```

**Example:**
```vb
http.AllowRedirects(False) _
    .SendGet("https://example.com/old-page")
```

## Manual Redirect Following

### FollowRedirect Method

```vb
Public Function FollowRedirect(Optional RedirectMethod As EnumRequestMethod = 0) As cHttpClient
```

**Basic Usage:**
```vb
http.AllowRedirects(False).SendGet("https://a.com")

If http.Inst.Status = 302 Then
    Debug.Print "Redirect to: " & http.GetRedirectUrl()
    http.FollowRedirect()  ' Auto request redirect URL
End If

Debug.Print "Final content: " & http.ReturnText()
```

**Specify 307/308 Redirect Method:**

When encountering 307/308 status codes, GET is used by default. To preserve original request method, specify via parameter:

```vb
http.AllowRedirects(False).SendPost("https://a.com", "data=value")

If http.Inst.Status = 307 Then
    ' 307 redirect, use POST to preserve method
    http.FollowRedirect(ReqPost)
End If
```

| Parameter Value | Description |
|----------------|-------------|
| `0` (default) | 307/308 use GET |
| `ReqPost` | 307/308 use POST |
| `ReqGet` | 307/308 use GET |
| `ReqPut` | 307/308 use PUT |
| `ReqDelete` | 307/308 use DELETE |
| `ReqOptions` | 307/308 use OPTIONS |

### Handle Redirect Chain

```vb
http.AllowRedirects(False).ResetRedirectCount()

Do
    http.SendGet(currentUrl)

    If http.Inst.Status >= 300 And http.Inst.Status < 400 Then
        Debug.Print "Redirect " & http.Inst.Status & " -> " & http.GetRedirectUrl()
        http.FollowRedirect()
    Else
        Exit Do
    End If
Loop
```

## Maximum Redirects

```vb
Public MaxRedirects As Long  ' Default 10 times
```

**Example:**
```vb
http.MaxRedirects = 5  ' Allow max 5 redirects

http.AllowRedirects(False).SendGet("https://a.com")
http.FollowRedirect()  ' Count +1
http.FollowRedirect()  ' Count +2
' Exceeding MaxRedirects will throw error
```

## Reset Redirect Count

```vb
Public Function ResetRedirectCount() As cHttpClient
```

**Example:**
```vb
http.AllowRedirects(False).ResetRedirectCount().SendGet("https://a.com")
```

## Get Redirect URL

```vb
Public Function GetRedirectUrl() As String
```

Extract redirect URL from `Location` response header:

```vb
Dim redirectUrl As String
redirectUrl = http.GetRedirectUrl()
If redirectUrl <> "" Then
    Debug.Print "Redirect to: " & redirectUrl
End If
```

## Redirect Behavior

| Status Code | Auto Redirect Behavior | Manual FollowRedirect Behavior |
|-------------|----------------------|-------------------------------|
| 301 | Auto to GET | Convert to GET |
| 302 | Auto to GET | Convert to GET |
| 303 | Auto to GET | Convert to GET |
| 307 | Preserve method | Default GET, can specify method |
| 308 | Preserve method | Default GET, can specify method |

## Practical Use Cases

### Scenario 1: Log Redirect Chain

```vb
Sub TraceRedirects()
    Dim http As New cHttpClient
    Dim redirectChain As String

    http.AllowRedirects(False).SendGet("https://bit.ly/3xxxxx")
    redirectChain = "Start URL"

    Do While http.Inst.Status >= 300 And http.Inst.Status < 400
        redirectChain = redirectChain & " -> " & http.GetRedirectUrl()
        http.FollowRedirect()
    Loop

    Debug.Print "Redirect chain: " & redirectChain
    Debug.Print "Final status: " & http.Inst.Status
End Sub
```

### Scenario 2: Modify Redirect Request

```vb
' Add headers before redirect
http.AllowRedirects(False).SendGet("https://a.com")

If http.Inst.Status = 302 Then
    ' Add extra headers
    http.RequestHeaders.Add "X-Special-Header", "value"
    http.FollowRedirect()
End If
```

### Scenario 3: Detect Redirect Loop

```vb
http.AllowRedirects(False).ResetRedirectCount().SendGet("https://a.com")

Do While http.Inst.Status >= 300 And http.Inst.Status < 400
    If http.RedirectCount > 10 Then
        Debug.Print "Redirect loop detected!"
        Exit Do
    End If
    http.FollowRedirect()
Loop
```
