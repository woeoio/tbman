# cHttpClient - HTTP 客户端类

## 简介

`cHttpClient` 是基于 WinHTTP 组件封装的 HTTP 客户端类，提供简洁的链式调用 API，支持同步/异步请求、自动重定向、JSON/表单数据处理、文件上传下载等功能。

## 快速开始

```vb
Dim http As New cHttpClient

' 简单的 GET 请求
http.SendGet("https://api.example.com/data")
Debug.Print http.ReturnText()

' POST JSON 数据
http.SetRequestContentType(Json)
http.RequestDataJson.Add "name", "test"
http.SendPost("https://api.example.com/users")

' 文件下载
http.DownloadFile "https://example.com/file.pdf", "C:\Downloads\file.pdf"

' 文件上传
http.UploadFile "https://api.example.com/upload", "C:\Data\report.pdf"
```

## 主要特性

- ✅ **链式调用** - 流畅的 API 设计
- ✅ **同步/异步** - 支持两种请求模式
- ✅ **自动重定向** - 支持 3xx 状态码自动跟随
- ✅ **数据格式** - JSON、表单、纯文本等多种 Content-Type
- ✅ **文件操作** - 简单的文件上传和下载
- ✅ **编码支持** - UTF-8 编码自动处理
- ✅ **调试信息** - 完整的请求/响应日志

## 依赖组件

- Microsoft WinHTTP Services, version 5.1
- Microsoft Scripting Runtime
- cJson 类

## 详细文档

完整开发文档请参考：`Docs/en/httpclient/`

