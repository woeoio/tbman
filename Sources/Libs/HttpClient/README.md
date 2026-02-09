# cHttpClient - HTTP Client Class

## Introduction

`cHttpClient` is an HTTP client class wrapped around the WinHTTP component, providing a simple chainable API with support for synchronous/asynchronous requests, automatic redirects, JSON/form data processing, file upload/download, and more.

## Quick Start

```vb
Dim http As New cHttpClient

' Simple GET request
http.SendGet("https://api.example.com/data")
Debug.Print http.ReturnText()

' POST JSON data
http.SetRequestContentType(Json)
http.RequestDataJson.Add "name", "test"
http.SendPost("https://api.example.com/users")

' File download
http.DownloadFile "https://example.com/file.pdf", "C:\Downloads\file.pdf"

' File upload
http.UploadFile "https://api.example.com/upload", "C:\Data\report.pdf"
```

## Key Features

- ✅ **Method Chaining** - Fluent API design
- ✅ **Sync/Async** - Support for both request modes
- ✅ **Auto Redirect** - Automatic following of 3xx status codes
- ✅ **Data Formats** - JSON, form, plain text, and other Content-Type support
- ✅ **File Operations** - Simple file upload and download
- ✅ **Encoding** - Automatic UTF-8 encoding handling
- ✅ **Debug Info** - Complete request/response logging

## Dependencies

- Microsoft WinHTTP Services, version 5.1
- Microsoft Scripting Runtime
- cJson class

## Documentation

For complete developer documentation, please refer to: `Docs/en/httpclient/`
