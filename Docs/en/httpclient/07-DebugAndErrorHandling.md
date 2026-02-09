# Debugging and Error Handling

## Debug Mode

### Enable Debug

```vb
Public DebugStart As Boolean
```

**Example:**
```vb
Dim http As New cHttpClient
http.DebugStart = True  ' Enable debug mode

http.SendGet("https://api.example.com/data")

' Output debug info
Debug.Print http.DebugInfo.Encode(, , True)
```

### Debug Info Structure

`DebugInfo` is a `cJson` object with the following fields:

```json
{
  "Request": {
    "IsAsync": false,
    "Method": "GET",
    "Url": "https://api.example.com/data",
    "Body": "",
    "Headers": {...},
    "TimeOut": 5,
    "ChartSet": "utf-8"
  },
  "Response": {
    "Status": 200,
    "StatusText": "OK",
    "Headers": {...},
    "Content": "..."
  },
  "Error": {
    "Number": 0,
    "Source": "",
    "Description": ""
  }
}
```

## Error Handling

### LastError Property

```vb
Public LastError As String
```

Stores the description of the last error.

**Example:**
```vb
http.SendGet("https://api.example.com/data")
If http.LastError <> "" Then
    Debug.Print "Error: " & http.LastError
End If
```

### Exception Message Format

When 4xx/5xx errors occur, exception description format is:

```
StatusCode#StatusText#ResponseContentFirst1024Chars
```

**Example:**
```vb
On Error GoTo ErrorHandler

http.SendGet("https://api.example.com/not-found")
' Throws exception: 404#Not Found#{"error": "Resource not found"}

Exit Sub

ErrorHandler:
    ' Parse error info
    Dim parts() As String
    parts = Split(Err.Description, "#")

    If UBound(parts) >= 2 Then
        Debug.Print "HTTP status code: " & parts(0)
        Debug.Print "Status text: " & parts(1)
        Debug.Print "Response content: " & parts(2)
    Else
        Debug.Print "Error: " & Err.Description
    End If
```

## Common Error Codes

| Error Code | Meaning | Solution |
|------------|---------|----------|
| 500 | Request URL is empty | Check URL parameter |
| 900 | Request timeout | Increase RequestTimeOut |
| 310 | Redirect count exceeded | Check redirect chain or increase MaxRedirects |
| 4xx | Client error | Check request parameters and authentication |
| 5xx | Server error | Check server status |

## Timeout Errors

### Set Timeout

```vb
http.RequestTimeOut = 60  ' 60 seconds
```

### Timeout Handling

```vb
On Error GoTo ErrorHandler

http.RequestTimeOut = 5
http.SendGet("https://slow-api.example.com/data")

Exit Sub

ErrorHandler:
    If Err.Number = 900 Then
        Debug.Print "Request timeout, please try again later"
    Else
        Debug.Print "Other error: " & Err.Description
    End If
```

## Network Error Handling

```vb
Sub SafeRequest()
    Dim http As New cHttpClient

    On Error GoTo ErrorHandler

    http.DebugStart = True
    http.RequestTimeOut = 30

    http.SendGet("https://api.example.com/data")

    ' Handle response
    ProcessResponse http.ReturnJson()

    Exit Sub

ErrorHandler:
    Dim errorMsg As String

    Select Case Err.Number
        Case 900
            errorMsg = "Request timeout, check network connection"
        Case 500
            errorMsg = "URL cannot be empty"
        Case 404
            errorMsg = "Resource not found"
        Case 401, 403
            errorMsg = "Permission denied, check authentication"
        Case Else
            errorMsg = "Request failed: " & Err.Description
    End Select

    Debug.Print errorMsg

    ' Output detailed debug info
    If Not http.DebugInfo Is Nothing Then
        Debug.Print "Detailed debug info:"
        Debug.Print http.DebugInfo.Encode(, , True)
    End If
End Sub
```

## Logging

### Log Complete Request

```vb
Sub LogRequest()
    Dim http As New cHttpClient
    http.DebugStart = True

    http.SendGet("https://api.example.com/data")

    ' Save to log file
    Dim logFile As String
    logFile = "C:\Logs\http_" & Format(Now, "yyyymmdd_hhmmss") & ".json"

    Open logFile For Output As #1
    Print #1, http.DebugInfo.Encode(, , True)
    Close #1

    Debug.Print "Log saved to: " & logFile
End Sub
```

## Debugging Tips

### 1. Compare Request and Response

```vb
With http.DebugInfo
    Debug.Print "Request URL: " & .Item("Request").Item("Url")
    Debug.Print "Response status: " & .Item("Response").Item("Status")
End With
```

### 2. Check Request Headers

```vb
Dim headers As Object
Set headers = http.DebugInfo.Item("Request").Item("Headers")

Dim key As Variant
For Each key In headers.Keys
    Debug.Print key & ": " & headers(key)
Next
```

### 3. Verify Response Content

```vb
Dim responseText As String
responseText = http.DebugInfo.Item("Response").Item("Content")

If Len(responseText) > 1000 Then
    Debug.Print "Response content too long, truncated"
End If
```
