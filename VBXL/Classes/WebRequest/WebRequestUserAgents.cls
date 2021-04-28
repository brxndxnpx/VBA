Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Module Summary
'''     A class for preset User-Agent headers (read-only).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''' Summary:
'''     A read-only User-Agent header for Netscape &3.
Public Property Get Netscape() As String: Netscape = "Mozilla/3.0 (Win95; I)": End Property

''' Summary:
'''     A read-only User-Agent header for WinPhone8.1.
Public Property Get WinPhone() As String: WinPhone = "Mozilla/5.0 (Mobile; Windows Phone 8.1; Android 4.0; ARM; Trident/7.0; Touch; rv:11.0; IEMobile/11.0; NOKIA; Lumia 520) like iPhone OS 7_0_3 Mac OS X AppleWebKit/537 (KHTML, like Gecko) Mobile Safari/537": End Property

''' Summary:
'''     A read-only User-Agent header for Safari5 (Win7).
Public Property Get Safari5() As String: Safari5 = "Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US) AppleWebKit/533.21.1 (KHTML, like Gecko) Version/5.0.5 Safari/533.21.1": End Property

''' Summary:
'''     A read-only User-Agent header for Safari9 (Mac).
Public Property Get Safari9() As String: Safari9 = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11) AppleWebKit/601.1.56 (KHTML, like Gecko) Version/9.0 Safari/601.1.56": End Property

''' Summary:
'''     A read-only User-Agent header for iPad.
Public Property Get iPad() As String: iPad = "Mozilla/5.0 (iPad; CPU OS 8_3 like Mac OS X) AppleWebKit/600.1.4 (KHTML, like Gecko) Version/8.0 Mobile/12F5027d Safari/600.1.4": End Property

''' Summary:
'''     A read-only User-Agent header for iPhone6.
Public Property Get iPhone6() As String: iPhone6 = "Mozilla/5.0 (iPhone; CPU iPhone OS 8_3 like Mac OS X) AppleWebKit/600.1.4 (KHTML, like Gecko) Version/8.0 Mobile/12F70 Safari/600.1.4": End Property

''' Summary:
'''     A read-only User-Agent header for IE 11 (Win8.1).
Public Property Get IE() As String: IE = "Mozilla/5.0 (Windows NT 6.3; WOW64; Trident/7.0; rv:11.0) like Gecko": End Property

''' Summary:
'''     A read-only User-Agent header for Edge (Win10).
Public Property Get Edge() As String: Edge = "Mozilla/5.0 (Windows NT 10.0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2486.0 Safari/537.36 Edge/13.11082": End Property

''' Summary:
'''     A read-only User-Agent header for Opera.
Public Property Get Opera() As String: Opera = "Opera/9.80 (Windows NT 6.2; WOW64) Presto/2.12.388 Version/12.17": End Property

''' Summary:
'''     A read-only User-Agent header for Firefox 43.
Public Property Get Firefox_Windows() As String: Firefox_Windows = "Mozilla/5.0 (Windows NT 6.3; WOW64; rv:43.0) Gecko/20100101 Firefox/43.0": End Property

''' Summary:
'''     A read-only User-Agent header for Firefox Phone.
Public Property Get Firefox_Phone() As String: Firefox_Phone = "Mozilla/5.0 (Mobile; rv:18.0) Gecko/18.0 Firefox/18.0": End Property

''' Summary:
'''     A read-only User-Agent header for Firefox (Mac).
Public Property Get Firefox_Mac() As String: Firefox_Mac = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10.8; rv:24.0) Gecko/20100101 Firefox/24.0": End Property

''' Summary:
'''     A read-only User-Agent header for Chrome (Win).
Public Property Get Chrome_Windows() As String: Chrome_Windows = "Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/48.0.2564.48 Safari/537.36": End Property

''' Summary:
'''     A read-only User-Agent header for Chrome (Android).
Public Property Get Chrome_Android() As String: Chrome_Android = "Mozilla/5.0 (Linux; Android 5.1.1; Nexus 5 Build/LMY48B) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/43.0.2357.78 Mobile Safari/537.36": End Property

''' Summary:
'''     A read-only User-Agent header for ChromeBook.
Public Property Get ChromeBook() As String: ChromeBook = "Mozilla/5.0 (X11; CrOS x86_64 6680.52.0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2272.74 Safari/537.36": End Property

''' Summary:
'''     A read-only User-Agent header for GoogleBot Crawler.
Public Property Get GoogleBotCrawler() As String: GoogleBotCrawler = "Mozilla/5.0 (compatible; Googlebot/2.1; +http://www.google.com/bot.html)": End Property

''' Summary:
'''     A read-only User-Agent header for Postman.
Public Property Get Postman() As String: Postman = "PostmanRuntime/7.26.10": End Property

