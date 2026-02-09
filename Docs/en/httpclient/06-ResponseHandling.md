# Response Handling

## Get Response Text

### ReturnText Method

```vb
Public Function ReturnText(Optional IsUtf8 As Boolean = True, Optional IsConvert As Boolean) As String
```

| Parameter | Description | Default |
|-----------|-------------|---------|
| IsUtf8 | Use UTF-8 decoding | True |
| IsConvert | Use StrConv conversion (fix garbled text) | False |

**Example:**
```vb
' Default UTF-8 decoding
Dim text As String
text = http.ReturnText()

' System encoding
text = http.ReturnText(False)

' Force conversion (handle garbled text)
text = http.ReturnText(False, True)
```

## Get JSON Data

### ReturnJson Method

```vb
Public Function ReturnJson(Optional IsUtf8 As Boolean = True, Optional IsConvert As Boolean) As cJson
```

**Example:**
```vb
http.SendGet("https://api.example.com/users/1")

Dim json As cJson
Set json = http.ReturnJson()

Debug.Print json.Item("name")
Debug.Print json.Item("email")
```

## Get Binary Data

### ReturnBody Method

```vb
Public Function ReturnBody() As Byte()
```

**Example:**
```vb
Dim bytes() As Byte
bytes = http.ReturnBody()

' Save to file
Open "C:\data.bin" For Binary As #1
Put #1, , bytes
Close #1
```

## Get Stream Data

### ReturnStream Method

```vb
Public Function ReturnStream() As Variant
```

**Example:**
```vb
Dim stream As Variant
Set stream = http.ReturnStream()
```

## Response Headers

### ResponseHeaders Dictionary

```vb
Public ResponseHeaders As New Scripting.Dictionary
```

**Example:**
```vb
http.SendGet("https://api.example.com/data")

' Get all headers
Dim key As Variant
For Each key In http.ResponseHeaders.Keys
    Debug.Print key & ": " & http.ResponseHeaders(key)
Next

' Get specific header
Dim contentType As String
If http.ResponseHeaders.Exists("Content-Type") Then
    contentType = http.ResponseHeaders("Content-Type")
End If
```

### Get Set-Cookie

```vb
' WinHttp auto-handles cookies, but can get via response headers
If http.ResponseHeaders.Exists("Set-Cookie") Then
    Debug.Print http.ResponseHeaders("Set-Cookie")
End If
```

## Status Code Handling

### Get Status Info

```vb
http.SendGet("https://api.example.com/data")

Debug.Print "Status code: " & http.Inst.Status
Debug.Print "Status text: " & http.Inst.StatusText
```

### Status Code Reference

| Range | Meaning | Handling |
|-------|---------|----------|
| 200-299 | Success | Normal return |
| 300-399 | Redirect | Auto follow or manual handle |
| 400-499 | Client error | Throw exception |
| 500-599 | Server error | Throw exception |

### Error Response Handling

```vb
On Error GoTo ErrorHandler

http.SendGet("https://api.example.com/not-found")

' If 4xx/5xx, throws exception
Debug.Print http.ReturnText()

Exit Sub

ErrorHandler:
    ' Err.Number contains status code
    ' Err.Description contains "Status#StatusText#ResponseContentFragment"
    Dim parts() As String
    parts = Split(Err.Description, "#")

    Debug.Print "Status code: " & parts(0)
    Debug.Print "Status text: " & parts(1)
    Debug.Print "Response: " & parts(2)
```

## Cookie Management

### Cookies Dictionary

```vb
Public Cookies As New Scripting.Dictionary
```

**Example:**
```vb
' Set cookie
http.SetCookies("session=abc123; path=/")

' Read cookie (need manual parsing)
Dim cookieValue As String
If http.Cookies.Exists("session") Then
    cookieValue = http.Cookies("session")
End If
```

## Response Data Cache

### ResponseRaw Property

```vb
Public ResponseRaw As Variant
```

Caches raw response content.

## Complete Response Handling Example

```vb
Sub HandleResponse()
    Dim http As New cHttpClient

    http.SendGet("https://api.example.com/users/1")

    ' 1. Check status code
    Select Case http.Inst.Status
        Case 200
            Debug.Print "Request successful"
        Case 201
            Debug.Print "Created"
        Case 204
            Debug.Print "No content"
        Case Else
            Debug.Print "Other status: " & http.Inst.Status
    End Select

    ' 2. Handle Content-Type
    Dim contentType As String
    If http.ResponseHeaders.Exists("Content-Type") Then
        contentType = http.ResponseHeaders("Content-Type")

        If InStr(contentType, "application/json") > 0 Then
            ' JSON response
            Dim json As cJson
            Set json = http.ReturnJson()
            ProcessJsonResponse json
        ElseIf InStr(contentType, "text/") > 0 Then
            ' Text response
            ProcessTextResponse http.ReturnText()
        Else
            ' Binary response
            ProcessBinaryResponse http.ReturnBody()
        End If
    End If

    ' 3. Get cookie
    If http.ResponseHeaders.Exists("Set-Cookie") Then
        SaveCookies http.ResponseHeaders("Set-Cookie")
    End If
End Sub
```
