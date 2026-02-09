# Advanced Usage

## Batch Requests

### Execute Multiple Requests Sequentially

```vb
Sub BatchRequests()
    Dim http As New cHttpClient
    Dim urls As Variant
    urls = Array( _
        "https://api.example.com/users/1", _
        "https://api.example.com/users/2", _
        "https://api.example.com/users/3" _
    )

    Dim i As Long
    For i = LBound(urls) To UBound(urls)
        http.SendGet(urls(i))
        Debug.Print "Request " & i + 1 & ": " & http.Inst.Status

        If http.Inst.Status = 200 Then
            ProcessUser http.ReturnJson()
        End If
    Next
End Sub
```

### Concurrent Requests (Using Collection)

```vb
' Use multiple instances for concurrency
Sub ConcurrentRequests()
    Dim http1 As New cHttpClient
    Dim http2 As New cHttpClient
    Dim http3 As New cHttpClient

    ' Send async
    http1.Async(True).SendGet("https://api.example.com/a")
    http2.Async(True).SendGet("https://api.example.com/b")
    http3.Async(True).SendGet("https://api.example.com/c")

    ' Wait for all requests to complete...
End Sub
```

## File Upload

### Simple File Upload (Recommended)

Use built-in `UploadFile` method to simplify file upload:

```vb
Sub SimpleUpload()
    Dim http As New cHttpClient

    ' One-line file upload
    If http.UploadFileSimple("https://api.example.com/upload", _
                              "C:\Documents\report.pdf") Then
        Debug.Print "Upload successful!"
        Debug.Print "Server response: " & http.ReturnText()
    Else
        Debug.Print "Upload failed: " & http.LastError
    End If
End Sub
```

### Advanced File Upload

Support custom field names and additional form data:

```vb
Sub AdvancedUpload()
    Dim http As New cHttpClient
    Dim formData As New Scripting.Dictionary

    ' Add additional form data
    formData.Add "userId", "12345"
    formData.Add "category", "documents"

    ' Upload file
    http.UploadFile "https://api.example.com/upload", _
                    "C:\Documents\report.pdf", _
                    fieldName:="file", _
                    additionalFormData:=formData

    ' Handle response
    If http.Inst.Status = 200 Then
        Debug.Print "Upload successful!"
    End If
End Sub
```

### Chain Upload

```vb
Sub ChainUpload()
    Dim http As New cHttpClient
    Dim result As cJson

    ' Upload and directly get JSON response
    Set result = http.UploadFile("https://api.example.com/upload", _
                                  "C:\Data\doc.docx", _
                                  "document") _
                         .ReturnJson()

    Debug.Print "File ID: " & result.Item("id")
End Sub
```

### Traditional Upload (Manual Request Building)

For more control, traditional method still available:

```vb
Sub ManualUpload()
    Dim http As New cHttpClient

    ' Read file to byte array
    Dim Stream As New cToolsStream
    Dim fileBytes() As Byte
    Stream.LoadFileAsBinary "C:\data.pdf", fileBytes

    ' Set multipart form and send
    http.SetRequestContentType(FormMultipart)
    http.RequestHeaders.Add "Content-Disposition", _
        "form-data; name=""file""; filename=""data.pdf"""

    http.SendPost "https://api.example.com/upload", _
                  StrConv(fileBytes, vbUnicode)
End Sub
```

---

## File Download

### Simple File Download (Recommended)

Use built-in `DownloadFile` method to complete download in one line:

```vb
Sub SimpleDownload()
    Dim http As New cHttpClient

    ' Download file (auto overwrite existing)
    If http.DownloadFile("https://example.com/file.pdf", _
                          "C:\Downloads\file.pdf") Then
        Debug.Print "Download successful!"
    End If

    ' Don't overwrite existing file
    http.DownloadFile "https://example.com/file.pdf", _
                      "C:\Downloads\file.pdf", _
                      Overwrite:=False
End Sub
```

### Async File Download

Use async mode to avoid blocking when downloading large files:

```vb
Private WithEvents httpDownload As cHttpClient

Sub AsyncDownload()
    Set httpDownload = New cHttpClient

    ' Start async download
    httpDownload.DownloadFileAsync "https://example.com/large-file.zip", _
                                   "C:\Downloads\large-file.zip"

    Debug.Print "Download started..."
End Sub

Private Sub httpDownload_OnResponseFinished()
    ' Save file after download complete
    If httpDownload.FinishDownloadFile() Then
        Debug.Print "Async download complete!"
    Else
        Debug.Print "Download failed: " & httpDownload.LastError
    End If
End Sub
```

### Batch Download

```vb
Sub BatchDownload()
    Dim http As New cHttpClient
    Dim files As Variant
    Dim i As Long

    files = Array( _
        Array("https://example.com/file1.pdf", "C:\Downloads\file1.pdf"), _
        Array("https://example.com/file2.pdf", "C:\Downloads\file2.pdf") _
    )

    For i = LBound(files) To UBound(files)
        On Error Resume Next
        http.DownloadFile files(i)(0), files(i)(1), Overwrite:=True

        If Err.Number = 0 Then
            Debug.Print "Download successful: " & files(i)(1)
        Else
            Debug.Print "Download failed: " & files(i)(0)
        End If
        On Error GoTo 0
    Next
End Sub
```

### Traditional Download (Manual Save)

For more control, traditional method still available:

```vb
Sub ManualDownload()
    Dim http As New cHttpClient

    http.SendGet("https://example.com/file.pdf")

    If http.Inst.Status = 200 Then
        ' Use cToolsStream to save file
        Dim Stream As New cToolsStream
        Dim fileData() As Byte
        fileData = http.ReturnBody()
        Stream.SaveFileAsBinary "C:\Downloads\file.pdf", fileData

        Debug.Print "File saved"
    End If
End Sub
```

## OAuth Authentication Flow

### Get Access Token

```vb
Function GetAccessToken() As String
    Dim http As New cHttpClient

    http.SetRequestContentType(FormUrlEncoded)
    http.RequestDataForm.Add "grant_type", "client_credentials"
    http.RequestDataForm.Add "client_id", "your_client_id"
    http.RequestDataForm.Add "client_secret", "your_client_secret"

    http.SendPost("https://oauth.example.com/token")

    If http.Inst.Status = 200 Then
        Dim json As cJson
        Set json = http.ReturnJson()
        GetAccessToken = json.Item("access_token")
    End If
End Function
```

### Use Token to Access API

```vb
Sub CallApiWithToken()
    Dim http As New cHttpClient

    ' Set auth header
    http.RequestHeaders.Add "Authorization", "Bearer " & GetAccessToken()

    http.SendGet("https://api.example.com/protected-resource")

    If http.Inst.Status = 200 Then
        ProcessData http.ReturnJson()
    ElseIf http.Inst.Status = 401 Then
        ' Token expired, refresh token and retry
        RefreshToken
        CallApiWithToken
    End If
End Sub
```

## Session Persistence

### Use Same Instance to Maintain Session

```vb
Sub SessionExample()
    Dim http As New cHttpClient

    ' Login
    http.RequestDataForm.Add "username", "admin"
    http.RequestDataForm.Add "password", "123456"
    http.SendPost("https://api.example.com/login")

    ' WinHttp auto-handles Set-Cookie

    ' Subsequent requests maintain session
    http.SendGet("https://api.example.com/user/profile")
    Debug.Print http.ReturnText()

    ' Another request, session still valid
    http.SendGet("https://api.example.com/user/settings")
    Debug.Print http.ReturnText()
End Sub
```

## Pagination Handling

### Auto Traverse Paginated API

```vb
Sub FetchAllPages()
    Dim http As New cHttpClient
    Dim allData As New Collection

    Dim page As Long
    page = 1

    Do
        http.RequestDataQuery.RemoveAll
        http.RequestDataQuery.Add "page", page
        http.RequestDataQuery.Add "limit", 100

        http.SendGet("https://api.example.com/items")

        If http.Inst.Status <> 200 Then Exit Do

        Dim json As cJson
        Set json = http.ReturnJson()

        Dim items As Object
        Set items = json.Item("data")

        Dim i As Long
        For i = 0 To items.Count - 1
            allData.Add items(i)
        Next

        ' Check if more pages
        If items.Count < 100 Then Exit Do
        page = page + 1

    Loop

    Debug.Print "Total fetched: " & allData.Count & " items"
End Sub
```

## Proxy Settings

### Set Proxy Server

```vb
Sub WithProxy()
    Dim http As New cHttpClient

    ' WinHttp proxy settings (need direct Inst access)
    http.Inst.SetProxy 2, "http://proxy.example.com:8080"

    ' If proxy requires auth
    http.Inst.SetCredentials "proxy_user", "proxy_pass", _
        HTTPREQUEST_SETCREDENTIALS_FOR_PROXY

    http.SendGet("https://api.example.com/data")
End Sub
```

## Custom Timeout

### Fine-grained Timeout Control

```vb
Sub CustomTimeouts()
    Dim http As New cHttpClient

    ' Direct WinHttpRequest for more granular timeouts
    ' SetTimeouts(DNS resolve timeout, connect timeout, send timeout, receive timeout)
    http.Inst.SetTimeouts 10000, 10000, 60000, 60000

    http.RequestTimeOut = 60  ' Sync wait timeout
    http.SendGet("https://slow-api.example.com/data")
End Sub
```

## Response Content Handling

### Auto Handle by Content-Type

```vb
Sub AutoProcessResponse()
    Dim http As New cHttpClient

    http.SendGet("https://api.example.com/data")

    Dim contentType As String
    If http.ResponseHeaders.Exists("Content-Type") Then
        contentType = http.ResponseHeaders("Content-Type")
    End If

    Select Case True
        Case InStr(contentType, "application/json") > 0
            ProcessJson http.ReturnJson()

        Case InStr(contentType, "application/xml") > 0
            ProcessXml http.ReturnText()

        Case InStr(contentType, "text/html") > 0
            ProcessHtml http.ReturnText()

        Case InStr(contentType, "image/") > 0
            SaveImage http.ReturnBody()

        Case Else
            ProcessRaw http.ReturnBody()
    End Select
End Sub
```

## Retry Mechanism

### Auto Retry Failed Requests

```vb
Function RetryRequest(url As String, maxRetries As Long) As cJson
    Dim http As New cHttpClient
    Dim attempt As Long

    For attempt = 1 To maxRetries
        On Error Resume Next

        http.SendGet(url)

        If Err.Number = 0 And http.Inst.Status < 400 Then
            Set RetryRequest = http.ReturnJson()
            Exit Function
        End If

        On Error GoTo 0

        Debug.Print "Request failed, retry " & attempt & "..."
        Sleep 1000 * attempt  ' Exponential backoff
    Next

    Err.Raise vbObjectError + 1, , "Request still failed after " & maxRetries & " attempts"
End Function
```
