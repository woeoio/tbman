# Request Methods

## Supported HTTP Methods

### GET Request

```vb
Public Function SendGet(ByVal url As String, Optional Body As String) As cHttpClient
```

**Example:**
```vb
http.SendGet("https://api.example.com/users")
http.SendGet("https://api.example.com/users", "id=123")  ' With body
```

### POST Request

```vb
Public Function SendPost(ByVal url As String, Optional Body As String) As cHttpClient
```

**Example:**
```vb
' Direct body
http.SendPost("https://api.example.com/users", "name=John&age=30")

' Using form data
http.RequestDataForm.Add "name", "John"
http.RequestDataForm.Add "age", "30"
http.SendPost("https://api.example.com/users")

' Using JSON
http.SetRequestContentType(Json)
http.RequestDataJson.Add "name", "John"
http.SendPost("https://api.example.com/users")
```

### PUT Request

```vb
Public Function SendPut(ByVal url As String, Optional Body As String) As cHttpClient
```

**Example:**
```vb
http.SetRequestContentType(Json)
http.RequestDataJson.Add "name", "Updated Name"
http.SendPut("https://api.example.com/users/1")
```

### DELETE Request

```vb
Public Function SendDelete(ByVal url As String, Optional Body As String) As cHttpClient
```

**Example:**
```vb
http.SendDelete("https://api.example.com/users/1")
```

### OPTIONS Request

```vb
Public Function SendOptions(ByVal url As String, Optional Body As String) As cHttpClient
```

**Example:**
```vb
http.SendOptions("https://api.example.com/users")
' Check Allow field in ResponseHeaders for supported methods
```

### Generic Request Method

```vb
Public Function Send(Method As EnumRequestMethod, ByVal Url As String, Optional Body As String) As cHttpClient
```

**Example:**
```vb
http.Send(ReqPatch, "https://api.example.com/users/1", "{\"name\":\"test\"}")
```

## Method Aliases

`Fetch` is an alias for `Send`:

```vb
http.Fetch(ReqGet, "https://api.example.com/data")  ' Same as Send
```

## Adding URL Query Parameters

```vb
' Use RequestDataQuery to add query parameters
http.RequestDataQuery.Add "page", "1"
http.RequestDataQuery.Add "limit", "10"
http.SendGet("https://api.example.com/users")
' Final URL: https://api.example.com/users?page=1&limit=10
```
