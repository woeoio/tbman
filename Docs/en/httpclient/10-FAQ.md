# Frequently Asked Questions

## Basic Questions

### Q: How to set request headers?

**A:** Use `RequestHeaders` dictionary:

```vb
http.RequestHeaders.Add "Authorization", "Bearer token123"
http.RequestHeaders.Add "X-Custom-Header", "value"
```

### Q: How to handle Chinese characters?

**A:** Default uses UTF-8 encoding, no extra handling needed:

```vb
http.RequestDataForm.Add "name", "中文内容"
http.SendPost("https://api.example.com/submit")
```

### Q: How to send JSON data?

**A:** Set Content-Type to JSON and add data:

```vb
http.SetRequestContentType(Json)
http.RequestDataJson.Add "key", "value"
http.RequestDataJson.Add "number", 123
http.SendPost("https://api.example.com/api")
```

---

## Error Handling

### Q: What to do about request timeout?

**A:** Increase timeout:

```vb
http.RequestTimeOut = 60  ' 60 seconds
```

### Q: How to handle 404 errors?

**A:** Use error handling:

```vb
On Error Resume Next
http.SendGet("https://api.example.com/not-found")

If http.Inst.Status = 404 Then
    Debug.Print "Resource not found"
ElseIf Err.Number <> 0 Then
    Debug.Print "Error: " & Err.Description
End If
On Error GoTo 0
```

### Q: How to handle SSL certificate errors?

**A:** Internally auto-ignores SSL certificate errors, no extra handling needed.

---

## Redirect Questions

### Q: How to disable auto redirect?

**A:** Use `AllowRedirects`:

```vb
http.AllowRedirects(False).SendGet("https://example.com/redirect")
Debug.Print http.Inst.Status  ' 302
Debug.Print http.GetRedirectUrl()  ' Get redirect address
```

### Q: How to manually follow redirect?

**A:** Use `FollowRedirect`:

```vb
http.AllowRedirects(False).SendGet("https://a.com")
If http.Inst.Status = 302 Then
    http.FollowRedirect()  ' Auto request redirect address
End If
```

### Q: How to limit redirect count?

**A:** Set `MaxRedirects`:

```vb
http.MaxRedirects = 5
```

---

## Response Handling Questions

### Q: What to do about garbled response text?

**A:** Try different decoding methods:

```vb
' Method 1: UTF-8 (default)
text = http.ReturnText(True)

' Method 2: System encoding
text = http.ReturnText(False)

' Method 3: Force conversion
text = http.ReturnText(False, True)
```

### Q: How to get binary response?

**A:** Use `ReturnBody`:

```vb
Dim bytes() As Byte
bytes = http.ReturnBody()
```

### Q: How to get response headers?

**A:** Use `ResponseHeaders`:

```vb
If http.ResponseHeaders.Exists("Content-Type") Then
    Debug.Print http.ResponseHeaders("Content-Type")
End If
```

---

## Cookie and Session

### Q: How to set cookies?

**A:** Use `SetCookies`:

```vb
http.SetCookies("session=abc123; user=john")
```

### Q: How to maintain session?

**A:** Use same instance:

```vb
Dim http As New cHttpClient

' Login
http.SendPost("https://api.example.com/login")

' Subsequent requests auto maintain session
http.SendGet("https://api.example.com/profile")
```

---

## Performance Optimization

### Q: How to improve request speed?

**A:**

1. Reuse instance (maintain connection reuse)
2. Set appropriate timeout
3. Use async requests

### Q: How to send multiple requests concurrently?

**A:** Use multiple instances:

```vb
Dim http1 As New cHttpClient
Dim http2 As New cHttpClient

http1.Async(True).SendGet("https://api.a.com")
http2.Async(True).SendGet("https://api.b.com")
```

---

## Debugging Questions

### Q: How to view complete request and response?

**A:** Enable debug mode:

```vb
http.DebugStart = True
http.SendGet("https://api.example.com/data")
Debug.Print http.DebugInfo.Encode(, , True)
```

### Q: How to log requests?

**A:**

```vb
http.DebugStart = True
http.SendGet("https://api.example.com/data")

Open "C:\log.txt" For Output As #1
Print #1, http.DebugInfo.Encode(, , True)
Close #1
```

---

## Special Scenarios

### Q: How to upload files?

**A:** Refer to file upload examples in Advanced Usage.

### Q: How to download files?

**A:**

```vb
http.SendGet("https://example.com/file.pdf")
If http.Inst.Status = 200 Then
    Open "C:\file.pdf" For Binary As #1
    Put #1, , http.ReturnBody()
    Close #1
End If
```

### Q: How to handle paginated API?

**A:**

```vb
Dim page As Long
page = 1

Do
    http.RequestDataQuery.Add "page", page
    http.SendGet("https://api.example.com/items")

    ' Process data...

    page = page + 1
Loop While HasMorePages(http.ReturnJson())
```

---

## Known Limitations

1. **WinHTTP Limitations** - Some advanced features depend on WinHTTP component capabilities
2. **Character Encoding** - Mainly supports UTF-8, other encodings may need extra handling
3. **Large Files** - Very large file upload/download may need streaming
4. **WebSocket** - WebSocket protocol not supported

---

## Getting Help

- Check `Docs/en/httpclient/` for complete documentation
- Refer to code examples
- Check debug info
