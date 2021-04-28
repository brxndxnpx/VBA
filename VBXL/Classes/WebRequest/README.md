# WebRequest

A late-binding wrapper class to execute web requests using the `WinHttp.WinHttpRequest` Windows API.

**Required Classes**
- [WebRequestContentTypes.cls](/VBXL/Classes/WebRequest/WebRequestContentTypes.cls)
- [WebRequestUserAgents.cls](/VBXL/Classes/WebRequest/WebRequestUserAgents.cls)

Recommended to be used in conjunction with...
- [JsonConverter.cls](/VBXL/Classes/JsonConverter/JsonConverter.cls) class to parse JSON results.


## Properties

| Property                     | Type                                | Description                                        | Default Value                                                                                                                         |
|------------------------------|-------------------------------------|----------------------------------------------------|---------------------------------------------------------------------------------------------------------------------------------------|
| `UserAgent`                  | `String`                            | The User-Agent header.                             | <small>Chrome Windows:</small><br>`Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/48.0.2564.48 Safari/537.36` |
| `ContentType`                | `String`                            | The Content-Type header.                           | `application/x-www-form-urlencoded`                                                                                                   |
| `AcceptEncoding`             | `String`                            | The Accept-Encoding header.                        | `gzip, deflate, br`                                                                                                                   |
| `Accept`                     | `String`                            | The Accept header.                                 | `*/*`                                                                                                                                 |
| `Async` *                    | `String`                            | Whether or not to execute requests asynchronously. | `False`                                                                                                                               |
| `UserAgents` [`read-only`]   | `Object` (`WebRequestUserAgents`)   | A class for preset User-Agent headers.             |                                                                                                                                       |
| `ContentTypes` [`read-only`] | `Object` (`WebRequestContentTypes`) | A class for preset Content-Type headers.           |                                                                                                                                       |

- `Async`: VBA doesn't have an await keyword. The code will continue and exit.
    - Use this when you don't need to wait for a response.


## Methods & Functions

|                                                     | Type     | Description                                                                                             |
|-----------------------------------------------------|----------|---------------------------------------------------------------------------------------------------------|
| [`Send`](#send)                                     | `String` | Executes a web request.                                                                                 |
| [`AddAuthorizationHeader`](#addauthorizationheader) |          | Adds an Authorization header to the request.                                                            |
| [`AddHeader`](#addheader)                           |          |                                                                                                         |
| [`GetAllResponseHeaders`](#getallresponseheaders)   | `String` | Gets all the headers from the HTTP response.                                                            |
| [`GetResponseHeader`](#getresponseheader)           | `String` | Gets a header from the HTTP response.                                                                   |
| [`SetCredentials`](#setcredentials)                 |          | Sets credentials to be used with an HTTP server, whether it is a proxy server or an originating server. |
| [`SetClientCertificate`](#setclientcertificate)     |          | Selects a client certificate to send to a Secure Hypertext Transfer Protocol (HTTPS) server.            |
| [`SetAutoLogonPolicy`](#setautologonpolicy)         |          | Sets the current Automatic Logon Policy.                                                                |
| [`SetProxy`](#setproxy)                             |          | Specify proxy configuration.                                                                            |
| [`SetTimeouts`](#settimeouts)                       |          | Specify timeout settings (in milliseconds).                                                             |


---

### `Send`

Executes a web request.

**Parameters**
- `ByVal RequestMethod As String`
    - The request method, i.e. GET, POST, PUT, DELETE, etc.
- `ByVal URL As String`
    - The url to send the request to.
- `[Optional] ByVal RequestBody As String`
    - The data to send with the request.

**Returns**
- `String`: The response text of the web request.

---

### `AddAuthorizationHeader`

Adds an Authorization header to the request.

**Parameters**
- `ByVal Value As String`
    - The header's value.
- `ByVal TokenType As String`
    - The token type; Bearer, Basic, etc.

---

### `AddHeader`

Adds a header to the request.

**Parameters**
- `ByVal Key As String` 
    - The header's key.
- `ByVal Value As String` 
    - The header's value.

---

### `GetAllResponseHeaders`

Gets all the headers from the HTTP response.

**Returns**
- `String`: The headers. Each header on a line.

---

### `GetResponseHeader`

Gets a header from the HTTP response.

**Parameters**
- `ByVal Header As String`
    - The header's key.

**Returns**
- `String`: The header's value.

---

### `SetCredentials`

Sets credentials to be used with an HTTP server, whether it is a proxy server or an originating server.
- [MSDN Documentation](https://docs.microsoft.com/en-us/windows/win32/winhttp/iwinhttprequest-setcredentials)

**Parameters**
- `ByVal UserName As String`
    - Specifies the user name for authentication.
- `ByVal Password As String`
    - Specifies the password for authentication. This parameter is ignored if bstrUserName is NULL or missing.
- `ByVal Flags As HTTPREQUEST_SETCREDENTIALS_FLAGS (Long)` 
    - Specifies when IWinHttpRequest uses credentials.
    
    ```vb
    ' WinHttp.WinHttpRequest enumeration for the SetCredentials method (verbatim)
    Public Enum HTTPREQUEST_SETCREDENTIALS_FLAGS
        ' Credentials are passed to a server.
        HTTPREQUEST_SETCREDENTIALS_FOR_SERVER = 0 

        ' Credentials are passed to a proxy.
        HTTPREQUEST_SETCREDENTIALS_FOR_PROXY = 1 
    End Enum
    ```
---

### `SetClientCertificate`

Selects a client certificate to send to a Secure Hypertext Transfer Protocol (HTTPS) server.
- Call SetClientCertificate to select a certificate before calling Send to send the request.
- [MSDN Documentation](https://docs.microsoft.com/en-us/windows/win32/winhttp/iwinhttprequest-setclientcertificate)

**Parameters**
- `ByVal ClientCertificate As String`
    - Specifies the location, certificate store, and subject of a client certificate.

---

### `SetAutoLogonPolicy`

Sets the current Automatic Logon Policy.
- [MSDN Documentation](https://docs.microsoft.com/en-us/windows/win32/winhttp/iwinhttprequest-setautologonpolicy)

**Parameters**
- `ByVal AutoLogonPolicy As WinHttpRequestAutoLogonPolicy (Long)`
    - Specifies the current automatic logon policy.

    ```vb
    ' WinHttp.WinHttpRequest enumeration for the SetAutoLogonPolicy method (verbatim).
    Public Enum WinHttpRequestAutoLogonPolicy
        AutoLogonPolicy_Always = 0
        AutoLogonPolicy_OnlyIfBypassProxy = 1
        AutoLogonPolicy_Never = 2
    End Enum
    ```

---

### `SetProxy`

Specify proxy configuration.
- [MSDN Documentation](https://docs.microsoft.com/en-us/windows/win32/winhttp/iwinhttprequest-setproxy)

**Parameters**
- `ByVal ProxySetting As HTTPREQUEST_PROXY_SETTING (Long)`
    - The flags that control this method.

    ```vb
    ' WinHttp.WinHttpRequest enumeration for the SetProxy method (verbatim)
    Public Enum HTTPREQUEST_PROXY_SETTING
        ' Indicates that the proxy settings should be obtained from the registry. 
        '   This assumes that Proxycfg.exe has been run.
        '   If Proxycfg.exe has not been run and HTTPREQUEST_PROXYSETTING_PRECONFIG is specified, 
        '   then the behavior is equivalent to HTTPREQUEST_PROXYSETTING_DIRECT.
        HTTPREQUEST_PROXYSETTING_PRECONFIG = 0
        
        ' Indicates that all HTTP and HTTPS servers should be accessed directly. 
        '   Use this command if there is no proxy server.
        HTTPREQUEST_PROXYSETTING_DIRECT = 1
        
        ' When HTTPREQUEST_PROXYSETTING_PROXY is specified, varProxyServer should 
        '   be set to a proxy server string and varBypassList should be set 
        '   to a domain bypass list string. This proxy configuration applies 
        '   only to the current instance of the WinHttpRequest object.
        HTTPREQUEST_PROXYSETTING_PROXY = 2
        
        ' Default proxy setting. 
        '   Equivalent to HTTPREQUEST_PROXYSETTING_PRECONFIG.
        [HTTPREQUEST_PROXYSETTING_DEFAULT] = HTTPREQUEST_PROXYSETTING_PRECONFIG
    End Enum
    ```

- `ByVal ProxyServer As String`
    - Set to a proxy server string when ProxySetting equals `HTTPREQUEST_PROXYSETTING_PROXY`.
- `ByVal BypassList As String`
    - Set to a domain bypass list string when ProxySetting equals `HTTPREQUEST_PROXYSETTING_PROXY`.

---

### `SetTimeouts`

Specify timeout settings (in milliseconds).
- [MSDN Documentation](https://docs.microsoft.com/en-us/windows/win32/winhttp/iwinhttprequest-settimeouts)

**Parameters**
- `ByVal ResolveTimeout As Long`
    - Time-out value applied when resolving a host name to an IP address.
    - The default value is zero, meaning no time-out (infinite).
- `ByVal ConnectTimeout As Long`
    - Time-out value applied when establishing a communication socket with the target server.
    - The default value is 60,000 (60 seconds).
- `ByVal SendTimeout As Long`
    - Time-out value applied when sending an individual packet of request data on the communication socket to the target server.
    - The default value is 30,000 (30 seconds).
- `ByVal ReceiveTimeout As Long`
    - Time-out value applied when receiving a packet of response data from the target server.
    - The default value is 30,000 (30 seconds).


---

## Usage

```vb
Private Sub Demo()
    Dim Client As New WebRequest
    Dim Result As String
    
    Result = Client.Send("GET", "https://animechan.vercel.app/api/random")
    
    Debug.Print Result
End Sub
```

---

## Notes

### Early Binding & Events

You can use events with the `WinHttp.WinHttpRequest` object if you use early binding methods.
- A reference to `Microsoft WinHTTP Services, version 5.1` is required.
    - In the Visual Basic Editor, click Tools > References > check "Microsoft WinHTTP Services, version 5.1"
- Please note that other users may have issues using your project/workbook if you use early bind your objects.
    - Each user would have to add a reference to the libraries.

```vb
' You can use events with WinHttp.WinHttpRequest if you use early binding.

' Replace this (late binding):
Private Client As Object

' With this (early binding): 
'   A reference to Microsoft WinHTTP Services, version 5.1 is required:
Private WithEvents Client As WinHttp.WinHttpRequest
```

Replacing the `CreateObject()` line in the constructor is actually optional. 
- It will create the same object.

```vb
''' Summary:
'''     The constructor. Sets the default values.
Private Sub Class_Initialize()
    ' This is the same as...
    Set Client = CreateObject("WinHttp.WinHttpRequest.5.1")

    ' this...
    Set Client = New WinHttp.WinHttpRequest
End Sub
```


Once `WinHttp.WinHttpRequest` is explicitly declared, you can use the events:

```vb
''' Summary
'''     Triggered after the response has been sent.
Private Sub Client_OnResponseStart(ByVal Status As Long, ByVal ContentType As String)
    Debug.Print Status, ContentType
End Sub

''' Summary
'''     Triggered after the OnResponseStart event.
Private Sub Client_OnResponseDataAvailable(Data() As Byte)
    Debug.Print "OnResponseDataAvailable"
End Sub

''' Summary
'''     Triggered after the OnResponseDataAvailable event.
Private Sub Client_OnResponseFinished()
    Debug.Print "OnResponseFinished"
End Sub

''' Summary
'''     Triggered if an error occurs with the request.
Private Sub Client_OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
    Debug.Print ErrorNumber, ErrorDescription
End Sub
```