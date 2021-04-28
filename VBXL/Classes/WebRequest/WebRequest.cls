Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Module Summary
'''     A late-binding wrapper class to execute web requests using the WinHttp.WinHttpRequest Windows API.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Web request libraries in the Windows registry.
'''     Microsoft.XMLHTTP {ED8C108E-4349-11D2-91A4-00C04F7969E8}
'''     MSXML2.XMLHTTP {F6D90F16-9C73-11D3-B32E-00C04F990BB4}
'''     WinHttp.WinHttpRequest.5.1 {2087c2f4-2cef-4953-a8ab-66779b670495}
'''     MSXML2.ServerXMLHTTP {AFBA6B42-5692-48EA-8141-DC517DCF0EF1}
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''' TODO: Create a download option for web requests using URLDownloadToFileA
'#If VBA7 Then
'    Private Declare PtrSafe Function URLDownloadToFile Lib "UrlMon" Alias "URLDownloadToFileA" (ByVal pCaller As LongPtr, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As LongPtr, ByVal lpfnCB As LongPtr) As LongPtr
'#Else
'    Private Declare Function URLDownloadToFile Lib "UrlMon" Alias "URLDownloadToFileA" ( ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
'#End If
'

Private Const AppPID As String = "WinHttp.WinHttpRequest.5.1"


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Remarks: You can use events with WinHttp.WinHttpRequest if you use early binding.
'''     See event snippet at the bottom of the class module.
'''     * A reference to Microsoft WinHTTP Services, version 5.1 is required:
'''
'''     Replace the line below (late binding):
Private Client As Object

'''     With this (early binding):
'Private WithEvents Client As WinHttp.WinHttpRequest
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''' Summary:
'''     WinHttp.WinHttpRequest enumeration for the SetCredentials method (verbatim)
''' Reference:
'''     https://docs.microsoft.com/en-us/windows/win32/winhttp/iwinhttprequest-setcredentials
Public Enum HTTPREQUEST_SETCREDENTIALS_FLAGS
    ' Credentials are passed to a server.
    HTTPREQUEST_SETCREDENTIALS_FOR_SERVER = 0

    ' Credentials are passed to a proxy.
    HTTPREQUEST_SETCREDENTIALS_FOR_PROXY = 1
End Enum


''' Summary:
'''     WinHttp.WinHttpRequest enumeration for the SetProxy method (verbatim)
''' Reference:
'''     https://docs.microsoft.com/en-us/windows/win32/winhttp/iwinhttprequest-setproxy
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


''' Summary:
'''     WinHttp.WinHttpRequest enumeration for the SetAutoLogonPolicy method (verbatim)
''' Reference:
'''     https://docs.microsoft.com/en-us/windows/win32/winhttp/winhttprequestautologonpolicy
Public Enum WinHttpRequestAutoLogonPolicy
    AutoLogonPolicy_Always = 0
    AutoLogonPolicy_OnlyIfBypassProxy = 1
    AutoLogonPolicy_Never = 2
End Enum


''' Summary:
'''     A type to store backing property values and private variables.
Private Type TWebRequest
    UserAgent       As String
    ContentType     As String
    Accept          As String
    AcceptEncoding  As String
    Async           As Boolean

    ResponseText    As String
    ResponseBody    As Variant
    ResponseStream  As Variant

    UserAgents      As New WebRequestUserAgents
    ContentTypes    As New WebRequestContentTypes
End Type

Private this As TWebRequest

''' Summary:
'''     The User-Agent header.
Public Property Get UserAgent() As String: UserAgent = this.UserAgent: End Property
Public Property Let UserAgent(ByVal Value As String): this.UserAgent = Value: End Property

''' Summary:
'''     The Content-Type header.
'''     Defaults to "application/x-www-form-urlencoded".
Public Property Get ContentType() As String: ContentType = this.ContentType: End Property
Public Property Let ContentType(ByVal Value As String): this.ContentType = Value: End Property

''' Summary:
'''     The Accept-Encoding header.
'''     Defaults to "gzip, deflate, br".
Public Property Get AcceptEncoding() As String: AcceptEncoding = this.AcceptEncoding: End Property
Public Property Let AcceptEncoding(ByVal Value As String): this.AcceptEncoding = Value: End Property

''' Summary:
'''     The Accept header.
'''     Defaults to "*/*".
Public Property Get Accept() As String: Accept = this.Accept: End Property
Public Property Let Accept(ByVal Value As String): this.Accept = Value: End Property

''' Summary:
'''     Whether or not to execute requests asynchronously. False by default.
'''     VBA doesn't have an await keyword. The code will continue and exit.
'''     Use this when you don't need to wait for a response.
Public Property Get Async() As Boolean: Async = this.Async: End Property
Public Property Let Async(ByVal Value As Boolean): this.Async = Value: End Property

''' Summary:
'''     A class for read-only User-Agent headers.
Public Property Get UserAgents() As WebRequestUserAgents: Set UserAgents = this.UserAgents: End Property

''' Summary:
'''     A class for read-only Content-Type headers.
Public Property Get ContentTypes() As WebRequestContentTypes: Set ContentTypes = this.ContentTypes: End Property


''' Summary:
'''     The constructor. Sets the default values.
Private Sub Class_Initialize()
    Set Client = CreateObject(AppPID)

    this.UserAgent = UserAgents.Chrome_Windows
    this.ContentType = ContentTypes.ApplicationJson
    this.Accept = "*/*"
    this.AcceptEncoding = "gzip, deflate, br"
End Sub

''' Summary:
'''     The deconstructor. Clears the objects from memory and closes any open connections.
Private Sub Class_Terminate()
    Set Client = Nothing
    Set this.UserAgents = Nothing
    Set this.ContentTypes = Nothing
End Sub


''' Summary:
'''     Executes a web request.
''' Parameters:
'''     ByVal RequestMethod As String: The request method, i.e. GET, POST, PUT, DELETE, etc.
'''     ByVal URL As String: The url to send the request to.
'''     Optional ByVal RequestBody As String: The data to send with the request.
''' Returns:
'''     A String; The response text of the web request.
Public Function Send(ByVal RequestMethod As String, ByVal URL As String, Optional ByVal RequestBody As String)
    Client.Open RequestMethod, URL, this.Async
    
    If this.UserAgent <> vbNullString Then AddHeader "User-Agent", this.UserAgent
    If this.ContentType <> vbNullString Then AddHeader "Content-Type", this.ContentType
    If this.AcceptEncoding <> vbNullString Then AddHeader "Accept-Encoding", this.AcceptEncoding
    If this.Accept <> vbNullString Then AddHeader "Accept", this.Accept

    If RequestBody = vbNullString Then Client.Send Else Client.Send RequestBody
    
    this.ResponseText = Client.ResponseText
    this.ResponseBody = Client.ResponseBody
'    this.ResponseStream = Client.ResponseStream
    
    Send = Client.ResponseText
End Function


''' Summary:
'''     Adds an Authorization header to the request.
''' Parameters:
'''     ByVal Value As String: The header's value.
'''     ByVal TokenType As String: The token type; Bearer, Basic, etc.
Public Sub AddAuthorizationHeader(ByVal Value As String, Optional ByVal TokenType As String)
    If Trim(TokenType) = vbNullString Then
        AddHeader "Authorization", Value
    Else
        AddHeader TokenType & " Authorization", Value
    End If
End Sub


''' Summary:
'''     Adds a header to the request.
''' Parameters:
'''     ByVal Key As String: The header's key.
'''     ByVal Value As String: The header's value.
Public Sub AddHeader(ByVal Key As String, ByVal Value As String)
    Client.SetRequestHeader Key, Value
End Sub


''' Summary:
'''     Gets all the headers from the HTTP response.
''' Returns:
'''     A String; The headers. Each header on a line.
Public Function GetAllResponseHeaders() As String
    GetAllResponseHeaders = Client.GetAllResponseHeaders
End Function


''' Summary:
'''     Gets a header from the HTTP response.
''' Parameters:
'''     ByVal Header As String: The header's key.
''' Returns:
'''     A String; The header's value.
Public Function GetResponseHeader(ByVal Header As String) As String
    GetResponseHeader = Client.GetResponseHeader(Header)
End Function


''' Summary:
'''     Sets credentials to be used with an HTTP server, whether it is a proxy server or an originating server.
''' Parameters:
'''     ByVal UserName As String: Specifies the user name for authentication.
'''     ByVal Password As String: Specifies the password for authentication. This parameter is ignored if bstrUserName is NULL or missing.
'''     ByVal Flags As HTTPREQUEST_SETCREDENTIALS_FLAGS (Long): Specifies when IWinHttpRequest uses credentials.
'''         HTTPREQUEST_SETCREDENTIALS_FOR_SERVER (0): Credentials are passed to a server.
'''         HTTPREQUEST_SETCREDENTIALS_FOR_PROXY  (1): Credentials are passed to a proxy.
''' Reference:
'''     https://docs.microsoft.com/en-us/windows/win32/winhttp/iwinhttprequest-setcredentials
Public Sub SetCredentials(ByVal UserName As String, ByVal Password As String, Flags As HTTPREQUEST_SETCREDENTIALS_FLAGS)
    Client.SetCredentials UserName, Password, Flags
End Sub


''' Summary:
'''     Selects a client certificate to send to a Secure Hypertext Transfer Protocol (HTTPS) server.
'''         Call SetClientCertificate to select a certificate before calling Send to send the request.
''' Parameters:
'''     ByVal ClientCertificate As String: Specifies the location, certificate store, and subject of a client certificate.
''' Reference:
'''     https://docs.microsoft.com/en-us/windows/win32/winhttp/iwinhttprequest-setclientcertificate
Public Sub SetClientCertificate(ByVal ClientCertificate As String)
    Client.SetClientCertificate ClientCertificate
End Sub


''' Summary:
'''     Sets the current Automatic Logon Policy.
''' Parameters:
'''     ByVal AutoLogonPolicy As WinHttpRequestAutoLogonPolicy (Long): Specifies the current automatic logon policy.
''' Reference:
'''     https://docs.microsoft.com/en-us/windows/win32/winhttp/iwinhttprequest-setautologonpolicy
Sub SetAutoLogonPolicy(AutoLogonPolicy As WinHttpRequestAutoLogonPolicy)
    Client.SetAutoLogonPolicy AutoLogonPolicy
End Sub


''' Summary:
'''     Specify proxy configuration.
''' Parameters:
'''     ByVal ProxySetting As HTTPREQUEST_PROXY_SETTING (Long): The flags that control this method.
'''     ByVal ProxyServer As String: Set to a proxy server string when ProxySetting equals HTTPREQUEST_PROXYSETTING_PROXY.
'''     ByVal BypassList As String: Set to a domain bypass list string when ProxySetting equals HTTPREQUEST_PROXYSETTING_PROXY.
''' Reference:
'''     https://docs.microsoft.com/en-us/windows/win32/winhttp/iwinhttprequest-setproxy
Sub SetProxy(ByVal ProxySetting As HTTPREQUEST_PROXY_SETTING, Optional ByRef ProxyServer As Variant, Optional ByRef BypassList As Variant)
    Client.SetProxy ProxySetting, ProxyServer, BypassList
End Sub


''' Summary:
'''     Specify timeout settings (in milliseconds).
''' Parameters:
'''     ByVal ResolveTimeout As Long: Time-out value applied when resolving a host name to an IP address.
'''         The default value is zero, meaning no time-out (infinite).
'''     ByVal ConnectTimeout As Long: Time-out value applied when establishing a communication socket with the target server.
'''         The default value is 60,000 (60 seconds).
'''     ByVal SendTimeout As Long: Time-out value applied when sending an individual packet of request data on the communication socket to the target server.
'''         The default value is 30,000 (30 seconds).
'''     ByVal ReceiveTimeout As Long: Time-out value applied when receiving a packet of response data from the target server.
'''         The default value is 30,000 (30 seconds).
''' Reference:
'''     https://docs.microsoft.com/en-us/windows/win32/winhttp/iwinhttprequest-settimeouts
Sub SetTimeouts(ResolveTimeout As Long, ConnectTimeout As Long, SendTimeout As Long, ReceiveTimeout As Long)
    Client.SetTimeouts ResolveTimeout, ConnectTimeout, SendTimeout, ReceiveTimeout
End Sub
    


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Remarks: You can use events with WinHttp.WinHttpRequest if you use early binding.
'''     See the top of the class module.
'''     * A reference to Microsoft WinHTTP Services, version 5.1 is required:
'''
''' Events:


''' Summary
'''     Triggered after the response has been sent.
'Private Sub Client_OnResponseStart(ByVal Status As Long, ByVal ContentType As String)
'    Debug.Print Status, ContentType
'End Sub


''' Summary
'''     Triggered after the OnResponseStart event.
'Private Sub Client_OnResponseDataAvailable(Data() As Byte)
'    Debug.Print "OnResponseDataAvailable"
'End Sub


''' Summary
'''     Triggered after the OnResponseDataAvailable event.
'Private Sub Client_OnResponseFinished()
'    Debug.Print "OnResponseFinished"
'End Sub


''' Summary
'''     Triggered if an error occurs with the request.
'Private Sub Client_OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
'    Debug.Print ErrorNumber, ErrorDescription
'End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''






