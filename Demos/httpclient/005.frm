' ===================================================================
' cHttpClient 文件上传和下载功能演示
' 利用 cToolsStream 简化文件 IO 操作
' ===================================================================

' ---------------------------------------------------------------
' 示例 1: 文件下载（同步方式）
' ---------------------------------------------------------------
Sub Demo_DownloadFile()
    Dim http As New cHttpClient
    Dim savePath As String
    
    savePath = "C:\Downloads\example.pdf"
    
    ' 一行代码下载文件
    If http.DownloadFile("https://example.com/file.pdf", savePath, Overwrite:=True) Then
        Debug.Print "文件下载成功，保存到: " & savePath
    End If
    
    ' 如果不想覆盖已存在的文件
    If http.DownloadFile("https://example.com/file.pdf", savePath, Overwrite:=False) Then
        Debug.Print "文件下载成功"
    Else
        Debug.Print "文件已存在或下载失败"
    End If
End Sub

' ---------------------------------------------------------------
' 示例 2: 文件下载（异步方式）
' ---------------------------------------------------------------
Private WithEvents httpDownload As cHttpClient

Sub Demo_DownloadFileAsync()
    Set httpDownload = New cHttpClient
    
    ' 启动异步下载
    httpDownload.DownloadFileAsync "https://example.com/large-file.zip", _
                                   "C:\Downloads\large-file.zip", _
                                   Overwrite:=True
    
    Debug.Print "下载已启动，继续执行其他代码..."
End Sub

Private Sub httpDownload_OnResponseFinished()
    ' 下载完成后自动保存文件
    If httpDownload.FinishDownloadFile() Then
        Debug.Print "异步下载完成并成功保存"
    Else
        Debug.Print "下载失败: " & httpDownload.LastError
    End If
End Sub

' ---------------------------------------------------------------
' 示例 3: 文件上传（简单方式）
' ---------------------------------------------------------------
Sub Demo_UploadFile()
    Dim http As New cHttpClient
    Dim filePath As String
    
    filePath = "C:\Documents\report.pdf"
    
    ' 简单上传（默认字段名为 "file"）
    If http.UploadFileSimple("https://api.example.com/upload", filePath) Then
        Debug.Print "上传成功"
        Debug.Print "服务器响应: " & http.ReturnText()
    Else
        Debug.Print "上传失败: " & http.LastError
    End If
End Sub

' ---------------------------------------------------------------
' 示例 4: 文件上传（带额外表单数据）
' ---------------------------------------------------------------
Sub Demo_UploadFileWithData()
    Dim http As New cHttpClient
    Dim filePath As String
    Dim formData As New Scripting.Dictionary
    
    filePath = "C:\Documents\avatar.jpg"
    
    ' 添加额外的表单数据
    formData.Add "userId", "12345"
    formData.Add "category", "profile"
    formData.Add "description", "用户头像"
    
    ' 上传文件，指定字段名和额外数据
    http.UploadFile "https://api.example.com/upload", _
                    filePath, _
                    fieldName:="avatar", _
                    additionalFormData:=formData
    
    ' 处理响应
    If http.Inst.Status = 200 Then
        Dim json As cJson
        Set json = http.ReturnJson()
        Debug.Print "上传成功，文件URL: " & json.Item("url")
    Else
        Debug.Print "上传失败: " & http.Inst.Status
    End If
End Sub

' ---------------------------------------------------------------
' 示例 5: 文件上传（链式调用方式）
' ---------------------------------------------------------------
Sub Demo_UploadFileChain()
    Dim http As New cHttpClient
    Dim result As cJson
    
    ' 链式调用上传并处理响应
    Set result = http.UploadFile("https://api.example.com/upload", _
                                  "C:\Data\document.docx", _
                                  "document") _
                         .ReturnJson()
    
    Debug.Print "文件ID: " & result.Item("id")
    Debug.Print "文件URL: " & result.Item("url")
End Sub

' ---------------------------------------------------------------
' 示例 6: 批量下载多个文件
' ---------------------------------------------------------------
Sub Demo_BatchDownload()
    Dim http As New cHttpClient
    Dim files As Variant
    Dim i As Long
    
    ' 定义要下载的文件列表
    files = Array( _
        Array("https://example.com/file1.pdf", "C:\Downloads\file1.pdf"), _
        Array("https://example.com/file2.pdf", "C:\Downloads\file2.pdf"), _
        Array("https://example.com/file3.pdf", "C:\Downloads\file3.pdf") _
    )
    
    For i = LBound(files) To UBound(files)
        On Error Resume Next
        
        http.DownloadFile files(i)(0), files(i)(1), Overwrite:=True
        
        If Err.Number = 0 Then
            Debug.Print "下载成功: " & files(i)(1)
        Else
            Debug.Print "下载失败: " & files(i)(0) & " - " & Err.Description
        End If
        
        On Error GoTo 0
    Next
End Sub

' ---------------------------------------------------------------
' 示例 7: 带进度显示的文件下载
' ---------------------------------------------------------------
Sub Demo_DownloadWithProgress()
    Dim http As New cHttpClient
    Dim url As String, savePath As String
    
    url = "https://example.com/large-file.zip"
    savePath = "C:\Downloads\large-file.zip"
    
    Debug.Print "开始下载..."
    
    If http.DownloadFile(url, savePath, Overwrite:=True) Then
        Dim fileSize As Long
        fileSize = FileLen(savePath)
        Debug.Print "下载完成! 文件大小: " & Format(fileSize / 1024 / 1024, "0.00") & " MB"
    Else
        Debug.Print "下载失败"
    End If
End Sub
