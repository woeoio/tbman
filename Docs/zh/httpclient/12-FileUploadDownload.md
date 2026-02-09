# 文件上传和下载

`cHttpClient` 提供了简洁的文件上传和下载功能，利用 `cToolsStream` 类简化了文件 IO 操作。

---

## 文件下载

### 同步下载

使用 `DownloadFile` 方法可以一行代码完成文件下载和保存：

```vb
Dim http As New cHttpClient

' 基本用法
http.DownloadFile "https://example.com/file.pdf", "C:\Downloads\file.pdf"

' 不覆盖已存在的文件
http.DownloadFile "https://example.com/file.pdf", _
                  "C:\Downloads\file.pdf", _
                  Overwrite:=False
```

**参数说明：**

| 参数 | 类型 | 说明 | 默认值 |
|------|------|------|--------|
| `url` | String | 文件 URL 地址 | 必需 |
| `savePath` | String | 本地保存路径 | 必需 |
| `Overwrite` | Boolean | 是否覆盖已存在文件 | `True` |

**返回值：**
- `True` - 下载成功
- 失败时抛出错误

---

### 异步下载

对于大文件下载，可以使用异步方式避免阻塞界面：

```vb
Private WithEvents http As cHttpClient

Sub StartDownload()
    Set http = New cHttpClient
    
    ' 启动异步下载
    http.DownloadFileAsync "https://example.com/large-file.zip", _
                           "C:\Downloads\large-file.zip"
End Sub

Private Sub http_OnResponseFinished()
    ' 下载完成后保存文件
    If http.FinishDownloadFile() Then
        Debug.Print "下载完成!"
    End If
End Sub
```

---

## 文件上传

### 简单上传

使用 `UploadFileSimple` 方法快速上传文件：

```vb
Dim http As New cHttpClient

If http.UploadFileSimple("https://api.example.com/upload", _
                          "C:\Documents\report.pdf") Then
    Debug.Print "上传成功!"
Else
    Debug.Print "上传失败: " & http.LastError
End If
```

---

### 高级上传

使用 `UploadFile` 方法可以自定义上传参数，支持链式调用：

```vb
Dim http As New cHttpClient
Dim formData As New Scripting.Dictionary

' 添加额外的表单数据
formData.Add "userId", "12345"
formData.Add "category", "documents"

' 上传文件
http.UploadFile "https://api.example.com/upload", _
                "C:\Documents\report.pdf", _
                fieldName:="file", _
                additionalFormData:=formData

' 处理响应
Debug.Print http.ReturnText()
```

**参数说明：**

| 参数 | 类型 | 说明 | 默认值 |
|------|------|------|--------|
| `url` | String | 上传地址 | 必需 |
| `filePath` | String | 本地文件路径 | 必需 |
| `fieldName` | String | 表单字段名 | `"file"` |
| `additionalFormData` | Dictionary | 额外的表单数据 | `Nothing` |

**返回值：**
- 返回 `cHttpClient` 实例，支持链式调用

---

### 链式调用示例

```vb
Dim json As cJson
Set json = http.UploadFile("https://api.example.com/upload", _
                            "C:\Data\doc.docx", _
                            "document") _
                   .ReturnJson()

Debug.Print "文件ID: " & json.Item("id")
```

---

## 使用建议

### 1. 下载前检查 URL

```vb
If url = "" Then
    Debug.Print "URL 不能为空"
    Exit Sub
End If

http.DownloadFile url, savePath
```

### 2. 处理下载错误

```vb
On Error Resume Next
http.DownloadFile url, savePath

If Err.Number <> 0 Then
    Select Case Err.Number
        Case 58
            Debug.Print "文件已存在"
        Case 200
            Debug.Print "服务器返回错误"
        Case Else
            Debug.Print "下载失败: " & Err.Description
    End Select
End If
On Error GoTo 0
```

### 3. 上传前检查文件

```vb
If Dir(filePath) = "" Then
    Debug.Print "文件不存在: " & filePath
    Exit Sub
End If

http.UploadFileSimple url, filePath
```

---

## 内部实现

文件上传下载功能内部使用了 `cToolsStream` 类处理文件 IO：

- **下载**：通过 `cToolsStream.SaveFileAsBinary()` 保存响应内容
- **上传**：通过 `cToolsStream.LoadFileAsBinary()` 读取文件内容

这种设计的好处：
1. 简化了文件操作流程
2. 统一的错误处理机制
3. 支持二进制文件
4. 自动处理文件编码

---

## 完整示例

```vb
Sub UploadAndDownloadDemo()
    Dim http As New cHttpClient
    Dim uploadPath As String
    Dim downloadPath As String

    uploadPath = "C:\Temp\data.txt"
    downloadPath = "C:\Temp\downloaded.bin"
    Const downloadText As String = "C:\Temp\downloaded.txt"

    ' ========== 上传文件 ==========
    Debug.Print "正在上传文件..."

    If http.UploadFileSimple("https://httpbin.org/post", uploadPath) Then
        Debug.Print "上传成功!"

        ' 解析响应
        Dim json As cJson
        Set json = http.ReturnJson()
        Debug.Print "服务器响应: " & json.Encode
    Else
        Debug.Print "上传失败: " & http.LastError
        Exit Sub
    End If

    ' ========== 下载文件 ==========
    Debug.Print "正在下载文件..."

    On Error Resume Next
    http.DownloadFile "https://httpbin.org/bytes/1024", downloadPath

    If Err.Number = 0 Then
        Debug.Print "下载成功! 文件大小: " & FileLen(downloadPath) & " 字节"
    Else
        Debug.Print "下载失败: " & Err.Description
    End If
    On Error GoTo 0

    ' ============= 下载文本 ===============
    On Error Resume Next
    http.DownloadFile "https://httpbin.org/encoding/utf8", downloadText

    If Err.Number = 0 Then
        Debug.Print "下载成功! 文件大小: " & FileLen(downloadText) & " 字节"
    Else
        Debug.Print "下载失败: " & Err.Description
    End If
    On Error GoTo 0

End Sub
```
