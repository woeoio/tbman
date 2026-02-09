# cHttpClient Overview

## Introduction

`cHttpClient` is an HTTP client class wrapper based on Microsoft WinHTTP Services 5.1, designed for TwinBASIC/VBA environments. It provides a modern chainable API that simplifies HTTP request handling.

## Design Goals

1. **Simple and Easy to Use** - Reduce code with method chaining
2. **Feature Complete** - Cover common HTTP request scenarios
3. **Flexible and Controllable** - Support custom headers, cookies, timeouts, etc.
4. **Error Friendly** - Clear status code handling and error messages

## Core Features

| Feature | Description |
|---------|-------------|
| HTTP Methods | GET, POST, PUT, DELETE, OPTIONS |
| Sync/Async | Support synchronous wait and asynchronous callback |
| Data Formats | JSON, form, plain text, binary |
| Redirect Handling | Auto/manual 3xx redirects |
| Encoding | Automatic UTF-8 encoding/decoding |
| SSL/TLS | Automatic handling with certificate error ignoring |
| Debug Logging | Complete request/response logging |

## Use Cases

- REST API calls
- Web service integration
- Data scraping
- File upload/download
- OAuth authentication flows

## Version Info

- Current version: Based on tbman project
- Author: Deng Wei (QQ: 215879458)
- License: Free to distribute, attribution required
