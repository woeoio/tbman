### 2025-02-09 v1.0.28
- [*] cHttpClient status code handling improvements:
        - Changed status check from 'Status <> 200' to 'Status >= 400 And Status <= 599'
        - 2xx success and 3xx redirects are no longer treated as errors
- [+] cHttpClient redirect handling features:
        - EnableRedirects property (auto-follow 3xx redirects, default True)
        - MaxRedirects property (limit redirect count, default 10)
        - AllowRedirects() method (chainable redirect toggle)
        - GetRedirectUrl() method (extract Location header)
        - FollowRedirect() method (manual redirect with method preservation)
        - ResetRedirectCount() method (chainable counter reset)
- [+] Add documentation:
        - Sources/Libs/HttpClient/README.md
        - Docs/httpclient/ (11 comprehensive guides)

### 2025-02-09 v1.0.27
- [+] cHttpClient add file transfer methods:
        - DownloadFile (sync download)
        - DownloadFileAsync (async download)
        - FinishDownloadFile (complete async download)
        - UploadFile (multipart/form-data upload with extra form data support)
        - UploadFileSimple (simplified upload returning boolean)
- [+] cHttpClient add private helper:
        - ConcatByteArrays (byte array concatenation for multipart body building)
- [+] Add demo: Demos/httpclient/005.frm (file upload/download examples)
- [+] Add documentation: Docs/httpclient/11-文件上传下载.md

### 2025-12-26 v1.0.26
- [*] cHttpClient add methods:
        - Send
        - SendPost
        - SendGet
        - SendPut
        - SendDelete
        - SendOptions

### 2025-12-26 v1.0.25(21)
- [-] Delete unused references(msscript.oxc, msdialog)

### 2025-01-10 v1.0.20
- [*] cSubClass Hook(+byval hWnd)

### 2025-01-10 v1.0.19
- [+] cSubClass

### 2024-12-30 v1.0.18
- [*] cJson .Items() assign

### 2024-12-30 v1.0.17
- [*] cIni.LoadFrom add Dic.RemoveAll()

### 2024-12-30 v1.0.16
- [+] cCsv

### 2024-12-30 v1.0.15
- [*] fix cIni "ToolsFso.AutoMakeDir Path, True"
- [*] fix cCsv "ToolsFso.AutoMakeDir Path, True"

### 2024-12-30 v1.0.14
- [+] cIni.Section
- [*] make cIni.Root default

### 2024-12-30 v1.0.13
- [*] fix cIni.Root init

### 2024-12-28 v1.0.12
- [*] fix cCsv.LoadFrom return cCsv
- [+] Add cIni members

### 2024-12-28 v1.0.11
- [*] fix cTimer x64 errors

### 2024-12-28 v1.0.10
- [+] cTimer
- [+] cTimers

### 2024-12-18 v1.0.9
- [+] ToolsDialog.OpenBox

### 2024-12-07 v1.0.7
- [+] cHttpClient

### 2024-12-07 v1.0.0
- Publish