# File Upload and Download

`cHttpClient` provides simple file upload and download functionality, leveraging the `cToolsStream` class to simplify file IO operations.

---

## File Download

### Synchronous Download

Use the `DownloadFile` method to complete file download and save in one line:

```vb
Dim http As New cHttpClient

' Basic usage
http.DownloadFile "https://example.com/file.pdf", "C:\Downloads\file.pdf"

' Don't overwrite existing file
http.DownloadFile "https://example.com/file.pdf", _
                  "C:\Downloads\file.pdf", _
                  Overwrite:=False
```

**Parameter Description:**

| Parameter | Type | Description | Default |
|-----------|------|-------------|---------|
| `url` | String | File URL | Required |
| `savePath` | String | Local save path | Required |
| `Overwrite` | Boolean | Overwrite existing file | `True` |

**Return Value:**
- `True` - Download successful
- Throws error on failure

---

### Asynchronous Download

For large file downloads, use async mode to avoid blocking the UI:

```vb
Private WithEvents http As cHttpClient

Sub StartDownload()
    Set http = New cHttpClient

    ' Start async download
    http.DownloadFileAsync "https://example.com/large-file.zip", _
                           "C:\Downloads\large-file.zip"
End Sub

Private Sub http_OnResponseFinished()
    ' Save file after download complete
    If http.FinishDownloadFile() Then
        Debug.Print "Download complete!"
    End If
End Sub
```

---

## File Upload

### Simple Upload

Use `UploadFileSimple` method for quick file upload:

```vb
Dim http As New cHttpClient

If http.UploadFileSimple("https://api.example.com/upload", _
                          "C:\Documents\report.pdf") Then
    Debug.Print "Upload successful!"
Else
    Debug.Print "Upload failed: " & http.LastError
End If
```

---

### Advanced Upload

Use `UploadFile` method to customize upload parameters, supporting method chaining:

```vb
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
Debug.Print http.ReturnText()
```

**Parameter Description:**

| Parameter | Type | Description | Default |
|-----------|------|-------------|---------|
| `url` | String | Upload URL | Required |
| `filePath` | String | Local file path | Required |
| `fieldName` | String | Form field name | `"file"` |
| `additionalFormData` | Dictionary | Additional form data | `Nothing` |

**Return Value:**
- Returns `cHttpClient` instance, supports method chaining

---

### Method Chaining Example

```vb
Dim json As cJson
Set json = http.UploadFile("https://api.example.com/upload", _
                            "C:\Data\doc.docx", _
                            "document") _
                   .ReturnJson()

Debug.Print "File ID: " & json.Item("id")
```

---

## Usage Recommendations

### 1. Check URL Before Download

```vb
If url = "" Then
    Debug.Print "URL cannot be empty"
    Exit Sub
End If

http.DownloadFile url, savePath
```

### 2. Handle Download Errors

```vb
On Error Resume Next
http.DownloadFile url, savePath

If Err.Number <> 0 Then
    Select Case Err.Number
        Case 58
            Debug.Print "File already exists"
        Case 200
            Debug.Print "Server returned error"
        Case Else
            Debug.Print "Download failed: " & Err.Description
    End Select
End If
On Error GoTo 0
```

### 3. Check File Before Upload

```vb
If Dir(filePath) = "" Then
    Debug.Print "File not found: " & filePath
    Exit Sub
End If

http.UploadFileSimple url, filePath
```

---

## Internal Implementation

File upload/download functionality internally uses `cToolsStream` class for file IO:

- **Download**: Save response content via `cToolsStream.SaveFileAsBinary()`
- **Upload**: Read file content via `cToolsStream.LoadFileAsBinary()`

Benefits of this design:
1. Simplifies file operation flow
2. Unified error handling mechanism
3. Supports binary files
4. Auto handles file encoding

---

## Complete Example

```vb
Sub UploadAndDownloadDemo()
    Dim http As New cHttpClient
    Dim uploadPath As String
    Dim downloadPath As String

    uploadPath = "C:\Temp\data.txt"
    downloadPath = "C:\Temp\downloaded.bin"
    Const downloadText As String = "C:\Temp\downloaded.txt"

    ' ========== Upload File ==========
    Debug.Print "Uploading file..."

    If http.UploadFileSimple("https://httpbin.org/post", uploadPath) Then
        Debug.Print "Upload successful!"

        ' Parse response
        Dim json As cJson
        Set json = http.ReturnJson()
        Debug.Print "Server response: " & json.Encode
    Else
        Debug.Print "Upload failed: " & http.LastError
        Exit Sub
    End If

    ' ========== Download File ==========
    Debug.Print "Downloading file..."

    On Error Resume Next
    http.DownloadFile "https://httpbin.org/bytes/1024", downloadPath

    If Err.Number = 0 Then
        Debug.Print "Download successful! File size: " & FileLen(downloadPath) & " bytes"
    Else
        Debug.Print "Download failed: " & Err.Description
    End If
    On Error GoTo 0

    ' ============= Download text ===============
    On Error Resume Next
    http.DownloadFile "https://httpbin.org/encoding/utf8", downloadText

    If Err.Number = 0 Then
        Debug.Print "Download successful! File size: " & FileLen(downloadText) & " bytes"
    Else
        Debug.Print "Download failed: " & Err.Description
    End If
    On Error GoTo 0

End Sub
```
